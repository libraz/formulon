//
// C ABI tests for workbook protection, protected ranges and note visibility.
// Every case crosses the ABI in both directions, including a save and load.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
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
constexpr fm_status_t kAlreadyExists = static_cast<fm_status_t>(formulon::FormulonErrorCode::kAlreadyExists);

void Create(WorkbookGuard* wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb->handle), 0) << fm_get_last_error_message();
}

void RoundTrip(const WorkbookGuard& wb, WorkbookGuard* out) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &out->handle), 0) << fm_get_last_error_message();
}

fm_protected_range_t Range(const char* name, const fm_cell_rect_t* ranges, size_t range_count) {
  fm_protected_range_t r = FM_STRUCT_INIT(fm_protected_range_t);
  r.name = name;
  r.ranges = ranges;
  r.range_count = static_cast<uint32_t>(range_count);
  return r;
}

fm_protected_range_patch_t Patch() {
  return FM_STRUCT_INIT(fm_protected_range_patch_t);
}

// The ranges of `r` as space-separated A1 text.
// Formatting releases the strings a read record points into, so call this after reading them.
std::string RangesText(const WorkbookGuard& wb, const fm_protected_range_t& r) {
  const std::vector<fm_cell_rect_t> ranges(r.ranges, r.ranges + r.range_count);
  std::string out;
  for (const fm_cell_rect_t& range : ranges) {
    const char* text = nullptr;
    EXPECT_EQ(fm_workbook_format_reference(wb.handle, nullptr, &range, 0, &text), 0);
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
  int32_t present = 7;
  ASSERT_EQ(fm_workbook_get_workbook_protection(wb.handle, &present, &p), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(p.lock_structure, 0);
  EXPECT_EQ(p.workbook_spin_count, 0U);
  EXPECT_EQ(p.workbook_hash_value, nullptr);
  EXPECT_EQ(p.workbook_legacy_password_hash, nullptr);
  EXPECT_EQ(p.revisions_salt_value, nullptr);
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
  in.workbook_legacy_password_hash = "CF1A";
  in.workbook_algorithm_name = "SHA-512";
  in.workbook_hash_value = "aGFzaA==";
  in.workbook_salt_value = "c2FsdA==";
  in.revisions_legacy_password_hash = "8D2B";
  in.revisions_algorithm_name = "SHA-256";
  in.revisions_hash_value = "cmV2";
  in.revisions_salt_value = "c2Fs";
  ASSERT_EQ(fm_workbook_set_workbook_protection(wb.handle, &in), 0) << fm_get_last_error_message();

  WorkbookGuard loaded;
  RoundTrip(wb, &loaded);
  fm_workbook_protection_t out = FM_STRUCT_INIT(fm_workbook_protection_t);
  int32_t present = 0;
  ASSERT_EQ(fm_workbook_get_workbook_protection(loaded.handle, &present, &out), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(out.lock_structure, 1);
  EXPECT_EQ(out.lock_windows, 1);
  EXPECT_EQ(out.lock_revision, 1);
  EXPECT_EQ(out.workbook_spin_count, 100000U);
  EXPECT_EQ(out.revisions_spin_count, 50000U);
  EXPECT_STREQ(out.workbook_legacy_password_hash, "CF1A");
  EXPECT_STREQ(out.workbook_algorithm_name, "SHA-512");
  EXPECT_STREQ(out.workbook_hash_value, "aGFzaA==");
  EXPECT_STREQ(out.workbook_salt_value, "c2FsdA==");
  EXPECT_STREQ(out.revisions_legacy_password_hash, "8D2B");
  EXPECT_STREQ(out.revisions_algorithm_name, "SHA-256");
  EXPECT_STREQ(out.revisions_hash_value, "cmV2");
  EXPECT_STREQ(out.revisions_salt_value, "c2Fs");
}

TEST(FormulonCApiWorkbookProtectionAbsence, FreshWorkbookNotPresent) {
  WorkbookGuard wb;
  Create(&wb);
  fm_workbook_protection_t out = FM_STRUCT_INIT(fm_workbook_protection_t);
  int32_t present = 1;
  ASSERT_EQ(fm_workbook_get_workbook_protection(wb.handle, &present, &out), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(out.lock_structure, 0);
  EXPECT_EQ(out.workbook_hash_value, nullptr);
}

TEST(FormulonCApiWorkbookProtectionAbsence, AllLocksOffIsPresent) {
  WorkbookGuard wb;
  Create(&wb);
  fm_workbook_protection_t in = FM_STRUCT_INIT(fm_workbook_protection_t);
  ASSERT_EQ(fm_workbook_set_workbook_protection(wb.handle, &in), 0);
  fm_workbook_protection_t out = FM_STRUCT_INIT(fm_workbook_protection_t);
  int32_t present = 0;
  ASSERT_EQ(fm_workbook_get_workbook_protection(wb.handle, &present, &out), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(out.lock_structure, 0);
}

TEST(FormulonCApiWorkbookProtectionAbsence, SetNullRemoves) {
  WorkbookGuard wb;
  Create(&wb);
  fm_workbook_protection_t in = FM_STRUCT_INIT(fm_workbook_protection_t);
  in.lock_structure = 1;
  ASSERT_EQ(fm_workbook_set_workbook_protection(wb.handle, &in), 0);
  ASSERT_EQ(fm_workbook_set_workbook_protection(wb.handle, nullptr), 0);
  fm_workbook_protection_t out = FM_STRUCT_INIT(fm_workbook_protection_t);
  int32_t present = 1;
  ASSERT_EQ(fm_workbook_get_workbook_protection(wb.handle, &present, &out), 0);
  EXPECT_EQ(present, 0);
  // Removing an absent element is a no-op.
  EXPECT_EQ(fm_workbook_set_workbook_protection(wb.handle, nullptr), 0);
}

TEST(FormulonCApiWorkbookProtection, NullArgumentsAreRejected) {
  WorkbookGuard wb;
  Create(&wb);
  fm_workbook_protection_t p = FM_STRUCT_INIT(fm_workbook_protection_t);
  int32_t present = 0;
  EXPECT_NE(fm_workbook_get_workbook_protection(nullptr, &present, &p), 0);
  EXPECT_NE(fm_workbook_get_workbook_protection(wb.handle, nullptr, &p), 0);
  EXPECT_NE(fm_workbook_get_workbook_protection(wb.handle, &present, nullptr), 0);
  EXPECT_NE(fm_workbook_set_workbook_protection(nullptr, &p), 0);
}

TEST(FormulonCApiProtectedRanges, AddListUpdateRemove) {
  WorkbookGuard wb;
  Create(&wb);
  fm_protected_range_t inputs = Range("Inputs", kInputsRanges, 2);
  inputs.legacy_password_hash = "CC1A";
  const fm_cell_rect_t totals_range{9, 0, 11, 0};
  fm_protected_range_t totals = Range("Totals", &totals_range, 1);
  totals.algorithm_name = "SHA-512";
  totals.hash_value = "aA==";
  totals.salt_value = "cw==";
  totals.spin_count = 100000;
  totals.security_descriptor = "O:WDG:WD";
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &inputs, nullptr), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &totals, nullptr), 0) << fm_get_last_error_message();
  EXPECT_EQ(RangeCount(wb), 2U);

  WorkbookGuard loaded;
  RoundTrip(wb, &loaded);
  ASSERT_EQ(RangeCount(loaded), 2U);
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  ASSERT_EQ(fm_sheet_get_protected_range_at(loaded.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.name, "Inputs");
  EXPECT_STREQ(out.legacy_password_hash, "CC1A");
  EXPECT_EQ(RangesText(loaded, out), "B2:C3 E5");
  ASSERT_EQ(fm_sheet_get_protected_range_at(loaded.handle, 0, 1, &out), 0);
  EXPECT_STREQ(out.algorithm_name, "SHA-512");
  EXPECT_STREQ(out.hash_value, "aA==");
  EXPECT_STREQ(out.salt_value, "cw==");
  EXPECT_EQ(out.spin_count, 100000U);
  EXPECT_STREQ(out.security_descriptor, "O:WDG:WD");

  const fm_cell_rect_t d4{3, 3, 3, 3};
  fm_protected_range_patch_t renamed = Patch();
  renamed.name_op = FM_PATCH_OP_SET;
  renamed.name = "Renamed";
  renamed.ranges_op = FM_PATCH_OP_SET;
  renamed.ranges = &d4;
  renamed.range_count = 1;
  ASSERT_EQ(fm_sheet_update_protected_range(loaded.handle, 0, "Inputs", &renamed), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_protected_range_at(loaded.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.name, "Renamed");
  EXPECT_STREQ(out.legacy_password_hash, "CC1A");
  EXPECT_EQ(out.algorithm_name, nullptr);
  EXPECT_EQ(RangesText(loaded, out), "D4");

  ASSERT_EQ(fm_sheet_remove_protected_range(loaded.handle, 0, "Renamed"), 0);
  ASSERT_EQ(RangeCount(loaded), 1U);
  ASSERT_EQ(fm_sheet_get_protected_range_at(loaded.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.name, "Totals");
  ASSERT_EQ(fm_sheet_remove_protected_range(loaded.handle, 0, "Totals"), 0);
  EXPECT_EQ(RangeCount(loaded), 0U);
}

TEST(FormulonCApiProtectedRanges, NamesCollideUnderUnicodeCaseFolding) {
  WorkbookGuard wb;
  Create(&wb);
  const fm_cell_rect_t a1{0, 0, 0, 0};
  const fm_cell_rect_t b2{1, 1, 1, 1};
  const fm_protected_range_t budget = Range("Budget", &a1, 1);
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &budget, nullptr), 0);

  const fm_protected_range_t upper = Range("BUDGET", &b2, 1);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &upper, nullptr), kAlreadyExists);
  const fm_protected_range_t umlaut = Range("\xC3\x84rger", &b2, 1);
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &umlaut, nullptr), 0);
  const fm_protected_range_t folded = Range("\xC3\xA4rger", &a1, 1);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &folded, nullptr), kAlreadyExists);
  EXPECT_EQ(RangeCount(wb), 2U);

  // A rename into the folded spelling of another range is refused; a recase of its own name is not.
  fm_protected_range_patch_t clash = Patch();
  clash.name_op = FM_PATCH_OP_SET;
  clash.name = "budget";
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "\xC3\x84rger", &clash), kAlreadyExists);
  fm_protected_range_patch_t recase = Patch();
  recase.name_op = FM_PATCH_OP_SET;
  recase.name = "BUDGET";
  ASSERT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Budget", &recase), 0);
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  ASSERT_EQ(fm_sheet_get_protected_range_at(wb.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.name, "BUDGET");
}

TEST(FormulonCApiProtectedRanges, RangesAreReadUpdatedAndRemovedByName) {
  WorkbookGuard wb;
  Create(&wb);
  const fm_cell_rect_t a1{0, 0, 0, 0};
  const fm_cell_rect_t b2{1, 1, 1, 1};
  const fm_protected_range_t umlaut = Range("\xC3\x84rger", &a1, 1);
  const fm_protected_range_t budget = Range("Budget", &b2, 1);
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &umlaut, nullptr), 0);
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &budget, nullptr), 0);

  // A name matches under Unicode case folding, whatever its position.
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_protected_range(wb.handle, 0, "\xC3\xA4rger", &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(out.name, "\xC3\x84rger");
  ASSERT_EQ(fm_sheet_get_protected_range(wb.handle, 0, "BUDGET", &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(RangesText(wb, out), "B2");

  fm_protected_range_patch_t patch = Patch();
  patch.spin_count_op = FM_PATCH_OP_SET;
  patch.spin_count = 7;
  ASSERT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "budget", &patch), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_protected_range(wb.handle, 0, "Budget", &present, &out), 0);
  EXPECT_EQ(out.spin_count, 7U);

  ASSERT_EQ(fm_sheet_remove_protected_range(wb.handle, 0, "\xC3\xA4rger"), 0);
  EXPECT_EQ(RangeCount(wb), 1U);
  ASSERT_EQ(fm_sheet_get_protected_range_at(wb.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.name, "Budget");
}

TEST(FormulonCApiProtectedRanges, UnknownNameReadsAbsentAndWritesNotFound) {
  WorkbookGuard wb;
  Create(&wb);
  const fm_cell_rect_t a1{0, 0, 0, 0};
  const fm_protected_range_t budget = Range("Budget", &a1, 1);
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &budget, nullptr), 0);

  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  int32_t present = 1;
  ASSERT_EQ(fm_sheet_get_protected_range(wb.handle, 0, "Other", &present, &out), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(out.name, nullptr);
  const fm_protected_range_patch_t keep = Patch();
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Other", &keep), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_remove_protected_range(wb.handle, 0, "Other"), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_remove_protected_range(wb.handle, 0, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_protected_range(wb.handle, 0, nullptr, &present, &out), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_protected_range(wb.handle, 4, "Budget", &present, &out), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(RangeCount(wb), 1U);
}

TEST(FormulonCApiProtectedRanges, SqrefFollowsRowAndColumnEdits) {
  WorkbookGuard wb;
  Create(&wb);
  const fm_protected_range_t r = Range("Inputs", kInputsRanges, 2);
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &r, nullptr), 0);
  ASSERT_EQ(fm_sheet_insert_rows(wb.handle, 0, 0, 1), 0);
  ASSERT_EQ(fm_sheet_insert_columns(wb.handle, 0, 0, 0), 0);
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  ASSERT_EQ(fm_sheet_get_protected_range_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(RangesText(wb, out), "C4:D5 F7");
  // Deleting every row the range covers removes it.
  ASSERT_EQ(fm_sheet_delete_rows(wb.handle, 0, 3, 6), 0);
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
  ASSERT_EQ(fm_sheet_get_protected_range_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(RangesText(wb, out), "B2");

  // Adding another range rewrites the list from the stored entries.
  const fm_cell_rect_t d4{3, 3, 3, 3};
  const fm_protected_range_t other = Range("Other", &d4, 1);
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &other, nullptr), 0);
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

  // KEEP leaves the rectangles alone, the unparsed token included.
  fm_protected_range_patch_t rename = Patch();
  rename.name_op = FM_PATCH_OP_SET;
  rename.name = "Renamed";
  ASSERT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Odd", &rename), 0) << fm_get_last_error_message();
  WorkbookGuard loaded;
  RoundTrip(wb, &loaded);
  const auto ranges = formulon::io::read_protected_ranges(loaded.handle->workbook().sheet(0).raw_extensions());
  ASSERT_EQ(ranges.size(), 1U);
  EXPECT_EQ(ranges[0].name, "Renamed");
  EXPECT_EQ(ranges[0].sqref, "B2 bogus");

  // A SET of an empty list is refused on update, and NULL with zero on add.
  const fm_cell_rect_t sentinel{0, 0, 0, 0};
  fm_protected_range_patch_t empty = Patch();
  empty.ranges_op = FM_PATCH_OP_SET;
  empty.ranges = &sentinel;
  empty.range_count = 0;
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Renamed", &empty), kInvalidArgument);
  const fm_protected_range_t no_ranges = Range("Renamed", nullptr, 0);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &no_ranges, nullptr), kInvalidArgument);
}

TEST(FormulonCApiProtectedRanges, ArgumentsAreValidated) {
  WorkbookGuard wb;
  Create(&wb);
  const fm_cell_rect_t a1{0, 0, 0, 0};
  const fm_cell_rect_t b1{0, 1, 0, 1};
  const fm_cell_rect_t inverted{2, 0, 1, 0};
  const fm_cell_rect_t off_grid{0, 0, 0, 16384};
  const fm_protected_range_t good = Range("Inputs", &a1, 1);
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &good, nullptr), 0);
  const fm_protected_range_t duplicate = Range("Inputs", &b1, 1);
  const fm_protected_range_t unnamed = Range("", &b1, 1);
  const fm_protected_range_t inverted_range = Range("Other", &inverted, 1);
  const fm_protected_range_t off_grid_range = Range("Other", &off_grid, 1);
  const fm_protected_range_t no_ranges = Range("Other", nullptr, 0);
  const fm_protected_range_t null_ranges = Range("Other", nullptr, 1);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &duplicate, nullptr), kAlreadyExists);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &unnamed, nullptr), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &inverted_range, nullptr), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &off_grid_range, nullptr), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &no_ranges, nullptr), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &null_ranges, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 9, &good, nullptr), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_NE(fm_sheet_add_protected_range(wb.handle, 0, nullptr, nullptr), 0);
  // Updating an entry to its own name is not a duplicate.
  fm_protected_range_patch_t same_name = Patch();
  same_name.name_op = FM_PATCH_OP_SET;
  same_name.name = "Inputs";
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Inputs", &same_name), 0);
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Missing", &same_name), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_remove_protected_range(wb.handle, 0, "Missing"), FM_STATUS_NOT_FOUND);
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  EXPECT_EQ(fm_sheet_get_protected_range_at(wb.handle, 0, 5, &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  // The context names the index by its parameter name.
  EXPECT_NE(std::string(fm_get_last_error_context()).find("protected_range_index="), std::string::npos)
      << fm_get_last_error_context();
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_protected_range_count(wb.handle, 9, &count), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(RangeCount(wb), 1U);
}

namespace {
fm_status_t SetNote(fm_workbook_t* wb, uint32_t row, uint32_t col, const char* text) {
  fm_comment_t note = FM_STRUCT_INIT(fm_comment_t);
  note.author = "Alice";
  note.text = text;
  return fm_sheet_set_comment(wb, 0, row, col, &note);
}
}  // namespace

TEST(FormulonCApiNoteVisibility, DefaultsToHiddenAndRoundTrips) {
  WorkbookGuard wb;
  Create(&wb);
  ASSERT_EQ(SetNote(wb.handle, 1, 1, "shown"), 0);
  ASSERT_EQ(SetNote(wb.handle, 3, 3, "hidden"), 0);
  int32_t present = 0;
  int32_t visible = 7;
  ASSERT_EQ(fm_sheet_get_comment_visible(wb.handle, 0, 1, 1, &present, &visible), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(visible, 0);
  ASSERT_EQ(fm_sheet_set_comment_visible(wb.handle, 0, 1, 1, 1), 0);

  WorkbookGuard loaded;
  RoundTrip(wb, &loaded);
  ASSERT_EQ(fm_sheet_get_comment_visible(loaded.handle, 0, 1, 1, &present, &visible), 0);
  EXPECT_EQ(visible, 1);
  ASSERT_EQ(fm_sheet_get_comment_visible(loaded.handle, 0, 3, 3, &present, &visible), 0);
  EXPECT_EQ(visible, 0);

  // Replacing the text keeps the flag; hiding it again persists.
  ASSERT_EQ(SetNote(loaded.handle, 1, 1, "edited"), 0);
  ASSERT_EQ(fm_sheet_get_comment_visible(loaded.handle, 0, 1, 1, &present, &visible), 0);
  EXPECT_EQ(visible, 1);
  ASSERT_EQ(fm_sheet_set_comment_visible(loaded.handle, 0, 1, 1, 0), 0);
  WorkbookGuard again;
  RoundTrip(loaded, &again);
  ASSERT_EQ(fm_sheet_get_comment_visible(again.handle, 0, 1, 1, &present, &visible), 0);
  EXPECT_EQ(visible, 0);
}

TEST(FormulonCApiNoteVisibility, ArgumentsAreValidated) {
  WorkbookGuard wb;
  Create(&wb);
  int32_t present = 7;
  int32_t visible = 7;
  // No note at the cell: the read succeeds and reports the absence.
  EXPECT_EQ(fm_sheet_get_comment_visible(wb.handle, 0, 1, 1, &present, &visible), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(visible, 0);
  EXPECT_EQ(fm_sheet_set_comment_visible(wb.handle, 0, 1, 1, 1), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_get_comment_visible(wb.handle, 9, 1, 1, &present, &visible), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_comment_visible(wb.handle, 9, 1, 1, 1), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_NE(fm_sheet_get_comment_visible(wb.handle, 0, 1, 1, &present, nullptr), 0);
  EXPECT_NE(fm_sheet_get_comment_visible(wb.handle, 0, 1, 1, nullptr, &visible), 0);
  EXPECT_NE(fm_sheet_get_comment_visible(nullptr, 0, 1, 1, &present, &visible), 0);
}

namespace {

// One stored range "Inputs" over B2:C3 E5 with every string set, and a second range "Totals".
void AddPatchFixture(const WorkbookGuard& wb) {
  fm_protected_range_t inputs = Range("Inputs", kInputsRanges, 2);
  inputs.legacy_password_hash = "CC1A";
  inputs.algorithm_name = "SHA-512";
  inputs.hash_value = "aA==";
  inputs.salt_value = "cw==";
  inputs.spin_count = 100000;
  inputs.security_descriptor = "O:WDG:WD";
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &inputs, nullptr), 0) << fm_get_last_error_message();
  const fm_cell_rect_t a10{9, 0, 9, 0};
  const fm_protected_range_t totals = Range("Totals", &a10, 1);
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &totals, nullptr), 0) << fm_get_last_error_message();
}

// Every field of range 0 equals the fixture.
void ExpectFixtureUnchanged(const WorkbookGuard& wb) {
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  ASSERT_EQ(fm_sheet_get_protected_range_at(wb.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.name, "Inputs");
  EXPECT_STREQ(out.legacy_password_hash, "CC1A");
  EXPECT_STREQ(out.algorithm_name, "SHA-512");
  EXPECT_STREQ(out.hash_value, "aA==");
  EXPECT_STREQ(out.salt_value, "cw==");
  EXPECT_EQ(out.spin_count, 100000U);
  EXPECT_STREQ(out.security_descriptor, "O:WDG:WD");
  EXPECT_EQ(RangesText(wb, out), "B2:C3 E5");
}

}  // namespace

TEST(FormulonCApiProtectedRangePatch, AllKeepIsANoOp) {
  WorkbookGuard wb;
  Create(&wb);
  AddPatchFixture(wb);
  const fm_protected_range_patch_t keep = Patch();
  ASSERT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Inputs", &keep), 0) << fm_get_last_error_message();
  ExpectFixtureUnchanged(wb);
}

TEST(FormulonCApiProtectedRangePatch, OneSetFieldChangesOnlyThatField) {
  WorkbookGuard wb;
  Create(&wb);
  AddPatchFixture(wb);
  fm_protected_range_patch_t patch = Patch();
  patch.spin_count_op = FM_PATCH_OP_SET;
  patch.spin_count = 5;
  ASSERT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Inputs", &patch), 0) << fm_get_last_error_message();
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  ASSERT_EQ(fm_sheet_get_protected_range_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.spin_count, 5U);
  EXPECT_STREQ(out.name, "Inputs");
  EXPECT_STREQ(out.hash_value, "aA==");
  EXPECT_STREQ(out.security_descriptor, "O:WDG:WD");
  EXPECT_EQ(RangesText(wb, out), "B2:C3 E5");

  patch = Patch();
  patch.security_descriptor_op = FM_PATCH_OP_SET;
  patch.security_descriptor = "";
  ASSERT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Inputs", &patch), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_protected_range_at(wb.handle, 0, 0, &out), 0);
  // An empty SET string writes the attribute absent.
  EXPECT_EQ(out.security_descriptor, nullptr);
  EXPECT_STREQ(out.salt_value, "cw==");
}

TEST(FormulonCApiProtectedRangePatch, ClearRemovesOnlyTheClearedPasswordStrings) {
  WorkbookGuard wb;
  Create(&wb);
  AddPatchFixture(wb);
  fm_protected_range_patch_t clear = Patch();
  clear.legacy_password_hash_op = FM_PATCH_OP_CLEAR;
  clear.hash_value_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Inputs", &clear), 0) << fm_get_last_error_message();
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  ASSERT_EQ(fm_sheet_get_protected_range_at(wb.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.legacy_password_hash, nullptr);
  EXPECT_EQ(out.hash_value, nullptr);
  EXPECT_STREQ(out.algorithm_name, "SHA-512");
  EXPECT_STREQ(out.salt_value, "cw==");
}

TEST(FormulonCApiProtectedRangePatch, RejectedPatchesChangeNothing) {
  WorkbookGuard wb;
  Create(&wb);
  AddPatchFixture(wb);
  const fm_cell_rect_t inverted{2, 0, 1, 0};
  const fm_cell_rect_t c3{2, 2, 2, 2};
  struct Case {
    const char* what;
    fm_protected_range_patch_t patch;
    fm_status_t expected;
  };
  std::vector<Case> cases;
  const auto add = [&cases](const char* what, fm_status_t expected, auto&& fill) {
    fm_protected_range_patch_t p = Patch();
    // A valid SET rides along so a partial write would show.
    p.salt_value_op = FM_PATCH_OP_SET;
    p.salt_value = "bmV3";
    fill(p);
    cases.push_back({what, p, expected});
  };
  add("op outside 0..2", kInvalidArgument, [](fm_protected_range_patch_t& p) {
    p.hash_value_op = static_cast<fm_patch_op_t>(3);
    p.hash_value = "eA==";
  });
  add("clear on the name", kInvalidArgument, [](fm_protected_range_patch_t& p) { p.name_op = FM_PATCH_OP_CLEAR; });
  add("clear on the spin count", kInvalidArgument,
      [](fm_protected_range_patch_t& p) { p.spin_count_op = FM_PATCH_OP_CLEAR; });
  add("clear on the ranges", kInvalidArgument, [](fm_protected_range_patch_t& p) { p.ranges_op = FM_PATCH_OP_CLEAR; });
  add("set with a NULL string", FM_STATUS_INVALID_ARGUMENT, [](fm_protected_range_patch_t& p) {
    p.algorithm_name_op = FM_PATCH_OP_SET;
    p.algorithm_name = nullptr;
  });
  add("set an empty name", kInvalidArgument, [](fm_protected_range_patch_t& p) {
    p.name_op = FM_PATCH_OP_SET;
    p.name = "";
  });
  add("set a taken name", kAlreadyExists, [](fm_protected_range_patch_t& p) {
    p.name_op = FM_PATCH_OP_SET;
    p.name = "Totals";
  });
  add("set an inverted rectangle", kInvalidArgument, [&inverted](fm_protected_range_patch_t& p) {
    p.ranges_op = FM_PATCH_OP_SET;
    p.ranges = &inverted;
    p.range_count = 1;
  });
  add("set NULL ranges with a count", FM_STATUS_BINDING_NULL_POINTER, [](fm_protected_range_patch_t& p) {
    p.ranges_op = FM_PATCH_OP_SET;
    p.ranges = nullptr;
    p.range_count = 1;
  });
  add("set malformed UTF-8", FM_STATUS_BINDING_UTF8ENCODING_ERROR, [&c3](fm_protected_range_patch_t& p) {
    p.ranges_op = FM_PATCH_OP_SET;
    p.ranges = &c3;
    p.range_count = 1;
    p.name_op = FM_PATCH_OP_SET;
    p.name = "\xff";
  });
  for (const Case& c : cases) {
    EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Inputs", &c.patch), c.expected) << c.what;
    ExpectFixtureUnchanged(wb);
  }
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Inputs", nullptr), FM_STATUS_BINDING_NULL_POINTER);
  const fm_protected_range_patch_t keep = Patch();
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 9, "Inputs", &keep), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "Missing", &keep), FM_STATUS_NOT_FOUND);
}

TEST(FormulonCApiProtectedRangeAddReturn, WritesName) {
  WorkbookGuard wb;
  Create(&wb);
  const fm_protected_range_t r = Range("Inputs", kInputsRanges, 2);
  const char* name = nullptr;
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &r, &name), 0) << fm_get_last_error_message();
  ASSERT_NE(name, nullptr);
  EXPECT_STREQ(name, "Inputs");
  fm_protected_range_t back = FM_STRUCT_INIT(fm_protected_range_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_protected_range(wb.handle, 0, name, &present, &back), 0);
  EXPECT_EQ(present, 1);
}
