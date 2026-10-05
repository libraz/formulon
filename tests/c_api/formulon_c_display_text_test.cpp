//
// Stable C ABI tests for cell display text and ad-hoc value formatting.

#include <cstdint>
#include <limits>
#include <string>
#include <utility>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

fm_value_t Number(double v) {
  fm_value_t value{};
  value.kind = FM_VAL_NUMBER;
  value.u.number = v;
  return value;
}

/// Renders through `fm_workbook_format_value` and returns `(text, status)`.
std::pair<std::string, int32_t> Format(fm_workbook_t* wb, const fm_value_t& value, const char* code) {
  const char* text = nullptr;
  int32_t status = -1;
  EXPECT_EQ(fm_workbook_format_value(wb, &value, code, &text, &status), 0);
  return {text == nullptr ? std::string() : std::string(text), status};
}

}  // namespace

TEST(FormulonCApiDisplayText, CellUsesItsNumberFormat) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf xf{};
  xf.num_fmt_id = 2;  // built-in 0.00
  uint32_t xf_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_index), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.5), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, xf_index), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 1, 0, "plain"), 0);

  const char* text = nullptr;
  int32_t status = -1;
  ASSERT_EQ(fm_workbook_get_display_text(wb.handle, 0, 0, 0, &text, &status), 0);
  EXPECT_STREQ(text, "1.50");
  EXPECT_EQ(status, FM_DISPLAY_OK);
  ASSERT_EQ(fm_workbook_get_display_text(wb.handle, 0, 1, 0, &text, &status), 0);
  EXPECT_STREQ(text, "plain");
  ASSERT_EQ(fm_workbook_get_display_text(wb.handle, 0, 5, 5, &text, &status), 0);
  EXPECT_STREQ(text, "");
  EXPECT_EQ(status, FM_DISPLAY_OK);
}

TEST(FormulonCApiDisplayText, OutOfCalendarDateOverflows) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const auto rendered = Format(wb.handle, Number(-1.0), "yyyy/m/d");
  EXPECT_EQ(rendered.first, "########");
  EXPECT_EQ(rendered.second, FM_DISPLAY_OVERFLOW);
}

TEST(FormulonCApiDisplayText, FormatValueCoversEveryScalarKind) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(Format(wb.handle, Number(1234.5), "#,##0").first, "1,235");
  EXPECT_EQ(Format(wb.handle, Number(0.25), "").first, "0.25");

  fm_value_t boolean{};
  boolean.kind = FM_VAL_BOOL;
  boolean.u.boolean = 1;
  EXPECT_EQ(Format(wb.handle, boolean, "General").first, "TRUE");

  fm_value_t text{};
  text.kind = FM_VAL_TEXT;
  text.u.text = "abc";
  EXPECT_EQ(Format(wb.handle, text, "General").first, "abc");

  fm_value_t error{};
  error.kind = FM_VAL_ERROR;
  error.u.error_code = 1;  // #DIV/0!
  EXPECT_EQ(Format(wb.handle, error, "0.00").first, "#DIV/0!");

  fm_value_t blank{};
  blank.kind = FM_VAL_BLANK;
  EXPECT_EQ(Format(wb.handle, blank, "0.00").first, "");
}

TEST(FormulonCApiDisplayText, FormatValueFollowsWorkbookDateSystem) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(Format(wb.handle, Number(1.0), "yyyy-mm-dd").first, "1900-01-01");
  wb.handle->workbook().set_date1904(true);
  EXPECT_EQ(Format(wb.handle, Number(1.0), "yyyy-mm-dd").first, "1904-01-02");
}

TEST(FormulonCApiDisplayText, RejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* text = nullptr;
  int32_t status = 0;
  EXPECT_EQ(fm_workbook_get_display_text(wb.handle, 3, 0, 0, &text, &status), kInvalidArgument);
  EXPECT_EQ(fm_workbook_get_display_text(wb.handle, 0, 0, 16384U, &text, &status), kInvalidArgument);
  EXPECT_NE(fm_workbook_get_display_text(wb.handle, 0, 0, 0, nullptr, &status), 0);

  fm_value_t value = Number(std::numeric_limits<double>::quiet_NaN());
  EXPECT_EQ(fm_workbook_format_value(wb.handle, &value, "0", &text, &status), kInvalidArgument);
  value.kind = FM_VAL_ARRAY;
  EXPECT_EQ(fm_workbook_format_value(wb.handle, &value, "0", &text, &status), kInvalidArgument);
  value.kind = FM_VAL_ERROR;
  value.u.error_code = 9999;
  EXPECT_EQ(fm_workbook_format_value(wb.handle, &value, "0", &text, &status), kInvalidArgument);
  value.kind = FM_VAL_TEXT;
  value.u.text = nullptr;
  EXPECT_NE(fm_workbook_format_value(wb.handle, &value, "0", &text, &status), 0);
  value = Number(1.0);
  EXPECT_NE(fm_workbook_format_value(wb.handle, &value, nullptr, &text, &status), 0);
  EXPECT_NE(fm_workbook_format_value(nullptr, &value, "0", &text, &status), 0);
}
