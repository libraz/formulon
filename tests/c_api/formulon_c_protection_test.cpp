// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Stable C ABI sheet-protection regression tests. Covers the
// `fm_sheet_get_protection` / `fm_sheet_set_protection` getter / setter
// pair, OOXML round-trip via `<sheetProtection>`, and NULL / range
// guards.

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

struct BufferGuard {
  uint8_t* data = nullptr;
  size_t len = 0;
  ~BufferGuard() { fm_buffer_free(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

}  // namespace

TEST(FormulonCApiProtection, FreshSheetReportsDisabled) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_sheet_protection_t p{};
  ASSERT_EQ(fm_sheet_get_protection(wb.handle, 0, &p), 0);
  EXPECT_EQ(p.enabled, 0);
  EXPECT_EQ(p.sheet, 0);
  EXPECT_EQ(p.objects, 0);
  EXPECT_EQ(p.spin_count, 0U);
  ASSERT_NE(p.algorithm_name, nullptr);
  EXPECT_STREQ(p.algorithm_name, "");
}

TEST(FormulonCApiProtection, SetThenGetReturnsAllFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_sheet_protection_t in{};
  in.enabled = 1;
  in.algorithm_name = "SHA-512";
  in.hash_value = "deadbeef==";
  in.salt_value = "saltsalt==";
  in.spin_count = 100000;
  in.legacy_password = "";  // modern path engaged
  in.sheet = 1;
  in.objects = 1;
  in.scenarios = 1;
  in.format_cells = 1;
  in.insert_columns = 1;
  in.select_locked_cells = 1;
  in.auto_filter = 1;
  ASSERT_EQ(fm_sheet_set_protection(wb.handle, 0, &in), 0);

  fm_sheet_protection_t out{};
  ASSERT_EQ(fm_sheet_get_protection(wb.handle, 0, &out), 0);
  EXPECT_EQ(out.enabled, 1);
  EXPECT_STREQ(out.algorithm_name, "SHA-512");
  EXPECT_STREQ(out.hash_value, "deadbeef==");
  EXPECT_STREQ(out.salt_value, "saltsalt==");
  EXPECT_EQ(out.spin_count, 100000U);
  EXPECT_EQ(out.sheet, 1);
  EXPECT_EQ(out.objects, 1);
  EXPECT_EQ(out.scenarios, 1);
  EXPECT_EQ(out.format_cells, 1);
  EXPECT_EQ(out.insert_columns, 1);
  EXPECT_EQ(out.select_locked_cells, 1);
  EXPECT_EQ(out.auto_filter, 1);
  // Fields the caller did not set should round-trip as 0.
  EXPECT_EQ(out.format_rows, 0);
  EXPECT_EQ(out.delete_columns, 0);
  EXPECT_EQ(out.pivot_tables, 0);
}

TEST(FormulonCApiProtection, NullStringsTreatedAsEmpty) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_sheet_protection_t in{};
  in.enabled = 1;
  in.sheet = 1;
  // All string pointers left as NULL.
  ASSERT_EQ(fm_sheet_set_protection(wb.handle, 0, &in), 0);

  fm_sheet_protection_t out{};
  ASSERT_EQ(fm_sheet_get_protection(wb.handle, 0, &out), 0);
  EXPECT_EQ(out.enabled, 1);
  ASSERT_NE(out.algorithm_name, nullptr);
  EXPECT_STREQ(out.algorithm_name, "");
  ASSERT_NE(out.hash_value, nullptr);
  EXPECT_STREQ(out.hash_value, "");
}

TEST(FormulonCApiProtection, OutOfRangeSheetReturnsInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_sheet_protection_t p{};
  EXPECT_EQ(fm_sheet_get_protection(wb.handle, 99, &p),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_sheet_set_protection(wb.handle, 99, &p),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiProtection, NullArgsReturnBindingNullPointer) {
  fm_sheet_protection_t p{};
  EXPECT_EQ(fm_sheet_get_protection(nullptr, 0, &p),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_set_protection(nullptr, 0, &p),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_get_protection(wb.handle, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_set_protection(wb.handle, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiProtection, ProtectedSheetRoundTripsThroughOoxml) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_sheet_protection_t in{};
  in.enabled = 1;
  in.algorithm_name = "SHA-512";
  in.hash_value = "AAAA==";
  in.salt_value = "BBBB==";
  in.spin_count = 100000;
  in.sheet = 1;
  in.objects = 1;
  in.format_cells = 1;
  in.select_locked_cells = 1;
  ASSERT_EQ(fm_sheet_set_protection(wb.handle, 0, &in), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=1"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);

  fm_sheet_protection_t out{};
  ASSERT_EQ(fm_sheet_get_protection(wb2.handle, 0, &out), 0);
  EXPECT_EQ(out.enabled, 1);
  EXPECT_STREQ(out.algorithm_name, "SHA-512");
  EXPECT_STREQ(out.hash_value, "AAAA==");
  EXPECT_STREQ(out.salt_value, "BBBB==");
  EXPECT_EQ(out.spin_count, 100000U);
  EXPECT_EQ(out.sheet, 1);
  EXPECT_EQ(out.objects, 1);
  EXPECT_EQ(out.format_cells, 1);
  EXPECT_EQ(out.select_locked_cells, 1);
  // Defaults preserved through round-trip.
  EXPECT_EQ(out.format_rows, 0);
  EXPECT_EQ(out.scenarios, 0);
}

TEST(FormulonCApiProtection, DisabledProtectionOmitsXmlElement) {
  // A workbook saved with `enabled = 0` must read back as still
  // disabled; the writer also must NOT emit a `<sheetProtection>`
  // element. We verify behavioural equivalence: a fresh load reports
  // enabled = 0.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // Stamp a non-default field value but leave `enabled = 0` to confirm
  // the writer respects the gate.
  fm_sheet_protection_t in{};
  in.enabled = 0;
  in.sheet = 1;
  in.objects = 1;
  ASSERT_EQ(fm_sheet_set_protection(wb.handle, 0, &in), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  fm_sheet_protection_t out{};
  ASSERT_EQ(fm_sheet_get_protection(wb2.handle, 0, &out), 0);
  EXPECT_EQ(out.enabled, 0);
  // Fields not persisted because the element was skipped.
  EXPECT_EQ(out.sheet, 0);
  EXPECT_EQ(out.objects, 0);
}

TEST(FormulonCApiProtection, LegacyPasswordRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_sheet_protection_t in{};
  in.enabled = 1;
  in.legacy_password = "CC53";
  in.sheet = 1;
  ASSERT_EQ(fm_sheet_set_protection(wb.handle, 0, &in), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  fm_sheet_protection_t out{};
  ASSERT_EQ(fm_sheet_get_protection(wb2.handle, 0, &out), 0);
  EXPECT_EQ(out.enabled, 1);
  EXPECT_STREQ(out.legacy_password, "CC53");
  EXPECT_EQ(out.sheet, 1);
}
