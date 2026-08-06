//
// Stable C ABI tests for the external-links surface
// (`fm_workbook_external_link_count` / `fm_workbook_external_link_at`).
// The reader/writer round-trip lives in the integration suite; these
// tests exercise the boundary's null / range guards and the empty-
// workbook behaviour.

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
