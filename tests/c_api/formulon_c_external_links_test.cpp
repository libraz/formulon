//
// Stable C ABI tests for the external-links surface
// (`fm_workbook_external_link_count` / `fm_workbook_external_link_at`).
// The reader/writer round-trip lives in the integration suite; these
// tests exercise the boundary's null / range guards and the empty-
// workbook behaviour.

#include <cstddef>
#include <cstdint>

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

uint32_t LinkCount(fm_workbook_t* wb) {
  uint32_t count = 99;
  EXPECT_EQ(fm_workbook_external_link_count(wb, &count), 0);
  return count;
}

}  // namespace

TEST(FormulonCApiExternalLinks, FreshWorkbookHasZeroLinks) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t count = 99;
  EXPECT_EQ(fm_workbook_external_link_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiExternalLinks, IndexOutOfRangeReturnsInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_external_link_record_t rec{};
  EXPECT_EQ(fm_workbook_external_link_at(wb.handle, 0, &rec),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiExternalLinks, NullArgsReturnBindingNullPointer) {
  fm_external_link_record_t rec{};
  uint32_t count = 0;
  EXPECT_EQ(fm_workbook_external_link_count(nullptr, &count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_external_link_at(nullptr, 0, &rec),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_external_link_count(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_external_link_at(wb.handle, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiExternalLinks, RoundTripPreservesZeroCount) {
  // No setter on the C ABI today; this guards the empty-default
  // round-trip the way the named-styles tests do. Reader-side records
  // surface through the integration suite.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint8_t* saved_data = nullptr;
  size_t saved_len = 0;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved_data, &saved_len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved_data, saved_len, &wb2.handle), 0);
  fm_buffer_free(saved_data);

  uint32_t count = 99;
  EXPECT_EQ(fm_workbook_external_link_count(wb2.handle, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiExternalLinks, SetFormulaBindsANewBookOnce) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=[Book.xlsx]Sheet1!A1"), 0) << fm_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 1U);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 1, 0, "=[BOOK.XLSX]Sheet1!B2"), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 2, 0, "=[book.xlsx]Other!A1"), 0) << fm_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 1U);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 3, 0, "=[Second.xlsx]Sheet1!A1"), 0) << fm_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 2U);
}

TEST(FormulonCApiExternalLinks, CfRuleNamingANewBookBindsIt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const fm_cf_cell_range_t sqref{0, 0, 9, 0};
  fm_cf_rule_t rule{};
  rule.type = 0;  // Expression
  rule.formula1 = "[Book.xlsx]Sheet1!$A$1>0";
  rule.sqref = &sqref;
  rule.sqref_count = 1;
  std::size_t index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &index), 0) << fm_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 1U);
  rule.formula1 = "[BOOK.xlsx]Sheet1!$A$2>0";
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &index), 0) << fm_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 1U);
  rule.formula1 = "A1>0";
  rule.formula2 = "[Other.xlsx]Sheet1!$A$1";
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, rule, &index), 0) << fm_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 2U);
}

TEST(FormulonCApiExternalLinks, ValidationNamingANewBookBindsIt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const fm_merge_range range{0, 0, 9, 0};
  fm_data_validation v{};
  v.ranges = &range;
  v.range_count = 1;
  v.type = 3;  // list
  v.formula1 = "[Book.xlsx]Sheet1!$A$1:$A$5";
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, v), 0) << fm_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 1U);
  v.formula1 = "[book.XLSX]Sheet1!$B$1:$B$5";
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, v), 0) << fm_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 1U);
  v.type = 1;
  v.formula1 = "0";
  v.formula2 = "[Other.xlsx]Sheet1!$A$1";
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, v), 0) << fm_last_error_message();
  EXPECT_EQ(LinkCount(wb.handle), 2U);
}
