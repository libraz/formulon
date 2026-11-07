//
// C ABI tests for workbook protection, protected ranges and note visibility.
// Every case crosses the ABI in both directions, including a save and load.

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "formulon_c_out_index.h"
#include "gtest/gtest.h"
#include "io/protection_xml.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

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
  ~BufferGuard() { fm_buffer_destroy(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNotFound = static_cast<fm_status_t>(formulon::FormulonErrorCode::kNotFound);

void Create(WorkbookGuard* wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb->handle), 0) << fm_get_last_error_message();
}

void RoundTrip(const WorkbookGuard& wb, WorkbookGuard* out) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &out->handle), 0) << fm_get_last_error_message();
}

fm_protected_range_t Range(const char* name, const fm_cell_rect_t* ranges, size_t range_count) {
  fm_protected_range_t r = FM_STRUCT_INIT(fm_protected_range_t);
  r.name = name;
  r.ranges = ranges;
  r.range_count = static_cast<uint32_t>(range_count);
  return r;
}

// The ranges of `r` as space-separated A1 text.
std::string RangesText(const WorkbookGuard& wb, const fm_protected_range_t& r) {
  std::string out;
  for (size_t i = 0; i < r.range_count; ++i) {
    const char* text = nullptr;
    EXPECT_EQ(fm_workbook_format_reference(wb.handle, nullptr, &r.ranges[i], 0, &text), 0);
    if (!out.empty()) {
      out.push_back(' ');
    }
    out += text != nullptr ? text : "";
  }
  return out;
}

// B2:C3 E5
const fm_cell_rect_t kInputsRanges[] = {{1, 1, 2, 2}, {4, 4, 4, 4}};

uint32_t RangeCount(const WorkbookGuard& wb) {
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_protected_range_count(wb.handle, 0, &count), 0);
  return static_cast<uint32_t>(count);
}

}  // namespace

TEST(FormulonCApiWorkbookProtection, AbsentByDefault) {
  WorkbookGuard wb;
  Create(&wb);
  fm_workbook_protection_t p = FM_STRUCT_INIT(fm_workbook_protection_t);
  p.lock_structure = 7;
  ASSERT_EQ(fm_workbook_get_workbook_protection(wb.handle, &p), 0);
  EXPECT_EQ(p.lock_structure, 0);
  EXPECT_EQ(p.workbook_spin_count, 0U);
  EXPECT_STREQ(p.workbook_hash_value, "");
}

TEST(FormulonCApiWorkbookProtection, EveryFieldRoundTripsThroughSaveAndLoad) {
  WorkbookGuard wb;
  Create(&wb);
  fm_workbook_protection_t in = FM_STRUCT_INIT(fm_workbook_protection_t);
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
  ASSERT_EQ(fm_workbook_set_workbook_protection(wb.handle, &in), 0) << fm_get_last_error_message();

  WorkbookGuard loaded;
  RoundTrip(wb, &loaded);
  fm_workbook_protection_t out = FM_STRUCT_INIT(fm_workbook_protection_t);
  ASSERT_EQ(fm_workbook_get_workbook_protection(loaded.handle, &out), 0);
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
  fm_workbook_protection_t in = FM_STRUCT_INIT(fm_workbook_protection_t);
  in.lock_structure = 1;
  ASSERT_EQ(fm_workbook_set_workbook_protection(wb.handle, &in), 0);
  fm_workbook_protection_t cleared = FM_STRUCT_INIT(fm_workbook_protection_t);
  ASSERT_EQ(fm_workbook_set_workbook_protection(wb.handle, &cleared), 0);
  WorkbookGuard loaded;
  RoundTrip(wb, &loaded);
  fm_workbook_protection_t out = FM_STRUCT_INIT(fm_workbook_protection_t);
  ASSERT_EQ(fm_workbook_get_workbook_protection(loaded.handle, &out), 0);
  EXPECT_EQ(out.lock_structure, 0);
}

TEST(FormulonCApiWorkbookProtection, NullArgumentsAreRejected) {
  WorkbookGuard wb;
  Create(&wb);
  fm_workbook_protection_t p = FM_STRUCT_INIT(fm_workbook_protection_t);
  EXPECT_NE(fm_workbook_get_workbook_protection(nullptr, &p), 0);
  EXPECT_NE(fm_workbook_get_workbook_protection(wb.handle, nullptr), 0);
  EXPECT_NE(fm_workbook_set_workbook_protection(nullptr, &p), 0);
  EXPECT_NE(fm_workbook_set_workbook_protection(wb.handle, nullptr), 0);
}

TEST(FormulonCApiProtectedRanges, AddListUpdateRemove) {
  WorkbookGuard wb;
  Create(&wb);
  fm_protected_range_t inputs = Range("Inputs", kInputsRanges, 2);
  inputs.password = "CC1A";
  const fm_cell_rect_t totals_range{9, 0, 11, 0};
  fm_protected_range_t totals = Range("Totals", &totals_range, 1);
  totals.algorithm_name = "SHA-512";
  totals.hash_value = "aA==";
  totals.salt_value = "cw==";
  totals.spin_count = 100000;
  totals.security_descriptor = "O:WDG:WD";
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &inputs, ignored_index()), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &totals, ignored_index()), 0) << fm_get_last_error_message();
  EXPECT_EQ(RangeCount(wb), 2U);

  WorkbookGuard loaded;
  RoundTrip(wb, &loaded);
  ASSERT_EQ(RangeCount(loaded), 2U);
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  ASSERT_EQ(fm_sheet_get_protected_range(loaded.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.name, "Inputs");
  EXPECT_EQ(RangesText(loaded, out), "B2:C3 E5");
  EXPECT_STREQ(out.password, "CC1A");
  ASSERT_EQ(fm_sheet_get_protected_range(loaded.handle, 0, 1, &out), 0);
  EXPECT_STREQ(out.algorithm_name, "SHA-512");
  EXPECT_STREQ(out.hash_value, "aA==");
  EXPECT_STREQ(out.salt_value, "cw==");
  EXPECT_EQ(out.spin_count, 100000U);
  EXPECT_STREQ(out.security_descriptor, "O:WDG:WD");

  const fm_cell_rect_t d4{3, 3, 3, 3};
  fm_protected_range_t renamed = Range("Renamed", &d4, 1);
  ASSERT_EQ(fm_sheet_update_protected_range(loaded.handle, 0, 0, &renamed), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_protected_range(loaded.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.name, "Renamed");
  EXPECT_EQ(RangesText(loaded, out), "D4");
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
  const fm_protected_range_t r = Range("Inputs", kInputsRanges, 2);
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &r, ignored_index()), 0);
  ASSERT_EQ(fm_sheet_insert_rows(wb.handle, 0, 0, 2), 0);
  ASSERT_EQ(fm_sheet_insert_columns(wb.handle, 0, 0, 1), 0);
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  ASSERT_EQ(fm_sheet_get_protected_range(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(RangesText(wb, out), "C4:D5 F7");
  // Deleting every row the range covers removes it.
  ASSERT_EQ(fm_sheet_delete_rows(wb.handle, 0, 3, 4), 0);
  EXPECT_EQ(RangeCount(wb), 0U);
}

TEST(FormulonCApiProtectedRanges, AnUnparsedTokenIsKeptForSave) {
  // A stored range the getter cannot express as a rectangle is left out of
  // `ranges`, but the stored text is what the file keeps.
  WorkbookGuard wb;
  Create(&wb);
  formulon::ProtectedRange stored;
  stored.name = "Odd";
  stored.sqref = "B2 bogus";
  formulon::io::write_protected_ranges(wb.handle->workbook().sheet(0).mutable_raw_extensions(), {stored});
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  ASSERT_EQ(fm_sheet_get_protected_range(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(RangesText(wb, out), "B2");

  // Adding another range rewrites the list from the stored entries.
  const fm_cell_rect_t d4{3, 3, 3, 3};
  const fm_protected_range_t other = Range("Other", &d4, 1);
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &other, ignored_index()), 0);
  WorkbookGuard loaded;
  RoundTrip(wb, &loaded);
  const auto ranges = formulon::io::read_protected_ranges(loaded.handle->workbook().sheet(0).raw_extensions());
  ASSERT_EQ(ranges.size(), 2U);
  EXPECT_EQ(ranges[0].sqref, "B2 bogus");
  EXPECT_EQ(ranges[1].sqref, "D4");
}

TEST(FormulonCApiProtectedRanges, UpdateWithoutRangesKeepsTheStoredText) {
  WorkbookGuard wb;
  Create(&wb);
  formulon::ProtectedRange stored;
  stored.name = "Odd";
  stored.sqref = "B2 bogus";
  formulon::io::write_protected_ranges(wb.handle->workbook().sheet(0).mutable_raw_extensions(), {stored});

  // NULL with a zero count leaves the rectangles alone, the unparsed token included.
  const fm_protected_range_t rename = Range("Renamed", nullptr, 0);
  ASSERT_EQ(fm_sheet_update_protected_range(wb.handle, 0, 0, &rename), 0) << fm_get_last_error_message();
  WorkbookGuard loaded;
  RoundTrip(wb, &loaded);
  const auto ranges = formulon::io::read_protected_ranges(loaded.handle->workbook().sheet(0).raw_extensions());
  ASSERT_EQ(ranges.size(), 1U);
  EXPECT_EQ(ranges[0].name, "Renamed");
  EXPECT_EQ(ranges[0].sqref, "B2 bogus");

  // An empty non-NULL list is still refused on update, and NULL with zero on add.
  const fm_cell_rect_t sentinel{0, 0, 0, 0};
  const fm_protected_range_t empty = Range("Empty", &sentinel, 0);
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, 0, &empty), kInvalidArgument);
  size_t index = 0;
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &rename, &index), kInvalidArgument);
}

TEST(FormulonCApiProtectedRanges, ArgumentsAreValidated) {
  WorkbookGuard wb;
  Create(&wb);
  const fm_cell_rect_t a1{0, 0, 0, 0};
  const fm_cell_rect_t b1{0, 1, 0, 1};
  const fm_cell_rect_t inverted{2, 0, 1, 0};
  const fm_cell_rect_t off_grid{0, 0, 0, 16384};
  const fm_protected_range_t good = Range("Inputs", &a1, 1);
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &good, ignored_index()), 0);
  const fm_protected_range_t duplicate = Range("Inputs", &b1, 1);
  const fm_protected_range_t unnamed = Range("", &b1, 1);
  const fm_protected_range_t inverted_range = Range("Other", &inverted, 1);
  const fm_protected_range_t off_grid_range = Range("Other", &off_grid, 1);
  const fm_protected_range_t no_ranges = Range("Other", nullptr, 0);
  const fm_protected_range_t null_ranges = Range("Other", nullptr, 1);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &duplicate, ignored_index()), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &unnamed, ignored_index()), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &inverted_range, ignored_index()), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &off_grid_range, ignored_index()), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &no_ranges, ignored_index()), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &null_ranges, ignored_index()), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 9, &good, ignored_index()), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_NE(fm_sheet_add_protected_range(wb.handle, 0, nullptr, ignored_index()), 0);
  // Updating an entry to its own name is not a duplicate.
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, 0, &good), 0);
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, 5, &good), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_remove_protected_range(wb.handle, 0, 5), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  EXPECT_EQ(fm_sheet_get_protected_range(wb.handle, 0, 5, &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_protected_range_count(wb.handle, 9, &count), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
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
  EXPECT_EQ(fm_sheet_get_comment_visible(wb.handle, 9, 1, 1, &visible), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_comment_visible(wb.handle, 9, 1, 1, 1), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_NE(fm_sheet_get_comment_visible(wb.handle, 0, 1, 1, nullptr), 0);
  EXPECT_NE(fm_sheet_get_comment_visible(nullptr, 0, 1, 1, &visible), 0);
}
