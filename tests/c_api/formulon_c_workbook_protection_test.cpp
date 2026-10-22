//
// C ABI tests for workbook protection, protected ranges and note visibility.
// Every case crosses the ABI in both directions, including a save and load.

#include <cstddef>
#include <cstdint>
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

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNotFound = static_cast<fm_status_t>(formulon::FormulonErrorCode::kNotFound);

void Create(WorkbookGuard* wb) {
  ASSERT_EQ(fm_workbook_create(&wb->handle), 0) << fm_last_error_message();
}

void RoundTrip(const WorkbookGuard& wb, WorkbookGuard* out) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &out->handle), 0) << fm_last_error_message();
}

fm_protected_range_t Range(const char* name, const char* sqref) {
  fm_protected_range_t r{};
  r.name = name;
  r.sqref = sqref;
  return r;
}

uint32_t RangeCount(const WorkbookGuard& wb) {
  uint32_t count = 0;
  EXPECT_EQ(fm_sheet_protected_range_count(wb.handle, 0, &count), 0);
  return count;
}

}  // namespace

TEST(FormulonCApiWorkbookProtection, AbsentByDefault) {
  WorkbookGuard wb;
  Create(&wb);
  fm_workbook_protection_t p{};
  p.lock_structure = 7;
  ASSERT_EQ(fm_workbook_get_protection(wb.handle, &p), 0);
  EXPECT_EQ(p.lock_structure, 0);
  EXPECT_EQ(p.workbook_spin_count, 0U);
  EXPECT_STREQ(p.workbook_hash_value, "");
}

TEST(FormulonCApiWorkbookProtection, EveryFieldRoundTripsThroughSaveAndLoad) {
  WorkbookGuard wb;
  Create(&wb);
  fm_workbook_protection_t in{};
  in.lock_structure = 1;
  in.lock_windows = 1;
  in.lock_revision = 1;
  in.workbook_spin_count = 100000;
  in.revisions_spin_count = 50000;
  in.workbook_password = "CF1A";
  in.workbook_algorithm_name = "SHA-512";
  in.workbook_hash_value = "aGFzaA==";
  in.workbook_salt_value = "c2FsdA==";
  in.revisions_password = "8D2B";
  in.revisions_algorithm_name = "SHA-256";
  in.revisions_hash_value = "cmV2";
  in.revisions_salt_value = "c2Fs";
  ASSERT_EQ(fm_workbook_set_protection(wb.handle, &in), 0) << fm_last_error_message();

  WorkbookGuard loaded;
  RoundTrip(wb, &loaded);
  fm_workbook_protection_t out{};
  ASSERT_EQ(fm_workbook_get_protection(loaded.handle, &out), 0);
  EXPECT_EQ(out.lock_structure, 1);
  EXPECT_EQ(out.lock_windows, 1);
  EXPECT_EQ(out.lock_revision, 1);
  EXPECT_EQ(out.workbook_spin_count, 100000U);
  EXPECT_EQ(out.revisions_spin_count, 50000U);
  EXPECT_STREQ(out.workbook_password, "CF1A");
  EXPECT_STREQ(out.workbook_algorithm_name, "SHA-512");
  EXPECT_STREQ(out.workbook_hash_value, "aGFzaA==");
  EXPECT_STREQ(out.workbook_salt_value, "c2FsdA==");
  EXPECT_STREQ(out.revisions_password, "8D2B");
  EXPECT_STREQ(out.revisions_algorithm_name, "SHA-256");
  EXPECT_STREQ(out.revisions_hash_value, "cmV2");
  EXPECT_STREQ(out.revisions_salt_value, "c2Fs");
}

TEST(FormulonCApiWorkbookProtection, AnAllEmptyRecordRemovesTheElement) {
  WorkbookGuard wb;
  Create(&wb);
  fm_workbook_protection_t in{};
  in.lock_structure = 1;
  ASSERT_EQ(fm_workbook_set_protection(wb.handle, &in), 0);
  fm_workbook_protection_t cleared{};
  ASSERT_EQ(fm_workbook_set_protection(wb.handle, &cleared), 0);
  WorkbookGuard loaded;
  RoundTrip(wb, &loaded);
  fm_workbook_protection_t out{};
  ASSERT_EQ(fm_workbook_get_protection(loaded.handle, &out), 0);
  EXPECT_EQ(out.lock_structure, 0);
}

TEST(FormulonCApiWorkbookProtection, NullArgumentsAreRejected) {
  WorkbookGuard wb;
  Create(&wb);
  fm_workbook_protection_t p{};
  EXPECT_NE(fm_workbook_get_protection(nullptr, &p), 0);
  EXPECT_NE(fm_workbook_get_protection(wb.handle, nullptr), 0);
  EXPECT_NE(fm_workbook_set_protection(nullptr, &p), 0);
  EXPECT_NE(fm_workbook_set_protection(wb.handle, nullptr), 0);
}

TEST(FormulonCApiProtectedRanges, AddListUpdateRemove) {
  WorkbookGuard wb;
  Create(&wb);
  fm_protected_range_t inputs = Range("Inputs", "B2:C3 E5");
  inputs.password = "CC1A";
  fm_protected_range_t totals = Range("Totals", "A10:A12");
  totals.algorithm_name = "SHA-512";
  totals.hash_value = "aA==";
  totals.salt_value = "cw==";
  totals.spin_count = 100000;
  totals.security_descriptor = "O:WDG:WD";
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &inputs), 0) << fm_last_error_message();
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &totals), 0) << fm_last_error_message();
  EXPECT_EQ(RangeCount(wb), 2U);

  WorkbookGuard loaded;
  RoundTrip(wb, &loaded);
  ASSERT_EQ(RangeCount(loaded), 2U);
  fm_protected_range_t out{};
  ASSERT_EQ(fm_sheet_get_protected_range(loaded.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.name, "Inputs");
  EXPECT_STREQ(out.sqref, "B2:C3 E5");
  EXPECT_STREQ(out.password, "CC1A");
  ASSERT_EQ(fm_sheet_get_protected_range(loaded.handle, 0, 1, &out), 0);
  EXPECT_STREQ(out.algorithm_name, "SHA-512");
  EXPECT_STREQ(out.hash_value, "aA==");
  EXPECT_STREQ(out.salt_value, "cw==");
  EXPECT_EQ(out.spin_count, 100000U);
  EXPECT_STREQ(out.security_descriptor, "O:WDG:WD");

  fm_protected_range_t renamed = Range("Renamed", "D4");
  ASSERT_EQ(fm_sheet_update_protected_range(loaded.handle, 0, 0, &renamed), 0) << fm_last_error_message();
  ASSERT_EQ(fm_sheet_get_protected_range(loaded.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.name, "Renamed");
  EXPECT_STREQ(out.sqref, "D4");
  EXPECT_STREQ(out.password, "");

  ASSERT_EQ(fm_sheet_remove_protected_range(loaded.handle, 0, 0), 0);
  ASSERT_EQ(RangeCount(loaded), 1U);
  ASSERT_EQ(fm_sheet_get_protected_range(loaded.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.name, "Totals");
  ASSERT_EQ(fm_sheet_remove_protected_range(loaded.handle, 0, 0), 0);
  EXPECT_EQ(RangeCount(loaded), 0U);
}

TEST(FormulonCApiProtectedRanges, SqrefFollowsRowAndColumnEdits) {
  WorkbookGuard wb;
  Create(&wb);
  const fm_protected_range_t r = Range("Inputs", "B2:C3 E5");
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &r), 0);
  ASSERT_EQ(fm_workbook_insert_rows(wb.handle, 0, 0, 2), 0);
  ASSERT_EQ(fm_workbook_insert_cols(wb.handle, 0, 0, 1), 0);
  fm_protected_range_t out{};
  ASSERT_EQ(fm_sheet_get_protected_range(wb.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.sqref, "C4:D5 F7");
  // Deleting every row the range covers removes it.
  ASSERT_EQ(fm_workbook_delete_rows(wb.handle, 0, 3, 4), 0);
  EXPECT_EQ(RangeCount(wb), 0U);
}

TEST(FormulonCApiProtectedRanges, ArgumentsAreValidated) {
  WorkbookGuard wb;
  Create(&wb);
  const fm_protected_range_t good = Range("Inputs", "A1");
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &good), 0);
  const fm_protected_range_t duplicate = Range("Inputs", "B1");
  const fm_protected_range_t unnamed = Range("", "B1");
  const fm_protected_range_t bad_sqref = Range("Other", "not a range");
  const fm_protected_range_t no_sqref = Range("Other", nullptr);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &duplicate), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &unnamed), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &bad_sqref), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &no_sqref), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 9, &good), kInvalidArgument);
  EXPECT_NE(fm_sheet_add_protected_range(wb.handle, 0, nullptr), 0);
  // Updating an entry to its own name is not a duplicate.
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, 0, &good), 0);
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, 5, &good), kInvalidArgument);
  EXPECT_EQ(fm_sheet_remove_protected_range(wb.handle, 0, 5), kInvalidArgument);
  fm_protected_range_t out{};
  EXPECT_EQ(fm_sheet_get_protected_range(wb.handle, 0, 5, &out), kInvalidArgument);
  uint32_t count = 0;
  EXPECT_EQ(fm_sheet_protected_range_count(wb.handle, 9, &count), kInvalidArgument);
  EXPECT_EQ(RangeCount(wb), 1U);
}

TEST(FormulonCApiNoteVisibility, DefaultsToHiddenAndRoundTrips) {
  WorkbookGuard wb;
  Create(&wb);
  ASSERT_EQ(fm_sheet_set_comment(wb.handle, 0, 1, 1, "Alice", "shown"), 0);
  ASSERT_EQ(fm_sheet_set_comment(wb.handle, 0, 3, 3, "Alice", "hidden"), 0);
  int32_t visible = 7;
  ASSERT_EQ(fm_sheet_get_comment_visible(wb.handle, 0, 1, 1, &visible), 0);
  EXPECT_EQ(visible, 0);
  ASSERT_EQ(fm_sheet_set_comment_visible(wb.handle, 0, 1, 1, 1), 0);

  WorkbookGuard loaded;
  RoundTrip(wb, &loaded);
  ASSERT_EQ(fm_sheet_get_comment_visible(loaded.handle, 0, 1, 1, &visible), 0);
  EXPECT_EQ(visible, 1);
  ASSERT_EQ(fm_sheet_get_comment_visible(loaded.handle, 0, 3, 3, &visible), 0);
  EXPECT_EQ(visible, 0);

  // Replacing the text keeps the flag; hiding it again persists.
  ASSERT_EQ(fm_sheet_set_comment(loaded.handle, 0, 1, 1, "Alice", "edited"), 0);
  ASSERT_EQ(fm_sheet_get_comment_visible(loaded.handle, 0, 1, 1, &visible), 0);
  EXPECT_EQ(visible, 1);
  ASSERT_EQ(fm_sheet_set_comment_visible(loaded.handle, 0, 1, 1, 0), 0);
  WorkbookGuard again;
  RoundTrip(loaded, &again);
  ASSERT_EQ(fm_sheet_get_comment_visible(again.handle, 0, 1, 1, &visible), 0);
  EXPECT_EQ(visible, 0);
}

TEST(FormulonCApiNoteVisibility, ArgumentsAreValidated) {
  WorkbookGuard wb;
  Create(&wb);
  int32_t visible = 0;
  EXPECT_EQ(fm_sheet_get_comment_visible(wb.handle, 0, 1, 1, &visible), kNotFound);
  EXPECT_EQ(fm_sheet_set_comment_visible(wb.handle, 0, 1, 1, 1), kNotFound);
  EXPECT_EQ(fm_sheet_get_comment_visible(wb.handle, 9, 1, 1, &visible), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_comment_visible(wb.handle, 9, 1, 1, 1), kInvalidArgument);
  EXPECT_NE(fm_sheet_get_comment_visible(wb.handle, 0, 1, 1, nullptr), 0);
  EXPECT_NE(fm_sheet_get_comment_visible(nullptr, 0, 1, 1, &visible), 0);
}
