//
// Stable C ABI sheet-protection regression tests. Covers the
// `fm_sheet_get_sheet_protection` / `fm_sheet_set_sheet_protection` /
// `fm_sheet_update_sheet_protection`, presence semantics, OOXML round-trip via
// `<sheetProtection>`, and NULL / range guards.

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
  ~BufferGuard() { fm_buffer_destroy(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

}  // namespace

namespace {

constexpr fm_status_t kInvalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

int32_t IsProtected(const fm_workbook_t* wb) {
  int32_t present = -1;
  fm_sheet_protection_t p = FM_STRUCT_INIT(fm_sheet_protection_t);
  EXPECT_EQ(fm_sheet_get_sheet_protection(wb, 0, &present, &p), 0);
  return present;
}

}  // namespace

TEST(FormulonCApiProtection, FreshSheetReportsNotProtected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  int32_t present = 7;
  fm_sheet_protection_t p = FM_STRUCT_INIT(fm_sheet_protection_t);
  ASSERT_EQ(fm_sheet_get_sheet_protection(wb.handle, 0, &present, &p), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(p.spin_count, 0U);
  EXPECT_EQ(p.allow_edit_objects, 0);
}

TEST(FormulonCApiProtection, SetThenGetReturnsAllFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_sheet_protection_t in = FM_STRUCT_INIT(fm_sheet_protection_t);
  in.algorithm_name = "SHA-512";
  in.hash_value = "deadbeef==";
  in.salt_value = "saltsalt==";
  in.spin_count = 100000;
  in.allow_edit_objects = 1;
  in.allow_edit_scenarios = 1;
  in.allow_format_cells = 1;
  in.allow_insert_columns = 1;
  in.allow_select_locked_cells = FM_TRISTATE_FALSE;
  in.allow_auto_filter = 1;
  ASSERT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &in), 0);

  int32_t present = 0;
  fm_sheet_protection_t out = FM_STRUCT_INIT(fm_sheet_protection_t);
  ASSERT_EQ(fm_sheet_get_sheet_protection(wb.handle, 0, &present, &out), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(out.algorithm_name, "SHA-512");
  EXPECT_STREQ(out.hash_value, "deadbeef==");
  EXPECT_STREQ(out.salt_value, "saltsalt==");
  EXPECT_EQ(out.spin_count, 100000U);
  EXPECT_EQ(out.allow_edit_objects, 1);
  EXPECT_EQ(out.allow_edit_scenarios, 1);
  EXPECT_EQ(out.allow_format_cells, 1);
  EXPECT_EQ(out.allow_insert_columns, 1);
  EXPECT_EQ(out.allow_select_locked_cells, FM_TRISTATE_FALSE);
  EXPECT_EQ(out.allow_auto_filter, 1);
  // Fields the caller did not set round-trip as not allowed.
  EXPECT_EQ(out.allow_format_rows, 0);
  EXPECT_EQ(out.allow_delete_columns, 0);
  EXPECT_EQ(out.allow_pivot_tables, 0);
}

TEST(FormulonCApiProtection, DefaultRecordAllowsSelectionAndNothingElse) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_protection_t in = FM_STRUCT_INIT(fm_sheet_protection_t);
  ASSERT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &in), 0);

  int32_t present = 0;
  fm_sheet_protection_t out = FM_STRUCT_INIT(fm_sheet_protection_t);
  ASSERT_EQ(fm_sheet_get_sheet_protection(wb.handle, 0, &present, &out), 0);
  EXPECT_EQ(present, 1);
  // The getter reports FALSE / TRUE, never DEFAULT.
  EXPECT_EQ(out.allow_select_locked_cells, FM_TRISTATE_TRUE);
  EXPECT_EQ(out.allow_select_unlocked_cells, FM_TRISTATE_TRUE);
  EXPECT_EQ(out.allow_format_cells, 0);
  EXPECT_EQ(out.allow_edit_objects, 0);
  EXPECT_EQ(out.allow_sort, 0);
}

TEST(FormulonCApiProtection, NullStringsReadAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_sheet_protection_t in = FM_STRUCT_INIT(fm_sheet_protection_t);
  // All string pointers left as NULL.
  ASSERT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &in), 0);

  int32_t present = 0;
  fm_sheet_protection_t out = FM_STRUCT_INIT(fm_sheet_protection_t);
  ASSERT_EQ(fm_sheet_get_sheet_protection(wb.handle, 0, &present, &out), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(out.algorithm_name, nullptr);
  EXPECT_EQ(out.hash_value, nullptr);
}

TEST(FormulonCApiProtection, SetWithNullRecordUnprotects) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_protection_t in = FM_STRUCT_INIT(fm_sheet_protection_t);
  in.allow_sort = 1;
  ASSERT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &in), 0);
  ASSERT_EQ(IsProtected(wb.handle), 1);

  EXPECT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, nullptr), 0);
  EXPECT_EQ(IsProtected(wb.handle), 0);
  // Unprotecting an unprotected sheet is a no-op, not an error.
  EXPECT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, nullptr), 0);
  EXPECT_EQ(IsProtected(wb.handle), 0);
}

TEST(FormulonCApiProtection, SetRejectsOutOfDomainFlagsWithoutChange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_protection_t bad_select = FM_STRUCT_INIT(fm_sheet_protection_t);
  bad_select.allow_select_locked_cells = 3;
  EXPECT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &bad_select), kInvalid);
  fm_sheet_protection_t bad_flag = FM_STRUCT_INIT(fm_sheet_protection_t);
  bad_flag.allow_sort = 2;
  EXPECT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &bad_flag), kInvalid);
  EXPECT_EQ(IsProtected(wb.handle), 0);
}

TEST(FormulonCApiProtection, UpdateOnUnprotectedSheetProtectsWithDefaults) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_protection_patch_t patch = FM_STRUCT_INIT(fm_sheet_protection_patch_t);
  patch.allow_sort_op = FM_PATCH_OP_SET;
  patch.allow_sort = 1;
  ASSERT_EQ(fm_sheet_update_sheet_protection(wb.handle, 0, &patch), 0);

  int32_t present = 0;
  fm_sheet_protection_t out = FM_STRUCT_INIT(fm_sheet_protection_t);
  ASSERT_EQ(fm_sheet_get_sheet_protection(wb.handle, 0, &present, &out), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(out.allow_sort, 1);
  EXPECT_EQ(out.allow_format_cells, 0);
  EXPECT_EQ(out.allow_select_locked_cells, FM_TRISTATE_TRUE);
}

TEST(FormulonCApiProtection, KeepOnlyUpdateLeavesUnprotectedSheetUnprotected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_protection_patch_t patch = FM_STRUCT_INIT(fm_sheet_protection_patch_t);
  EXPECT_EQ(fm_sheet_update_sheet_protection(wb.handle, 0, &patch), 0);
  EXPECT_EQ(IsProtected(wb.handle), 0);
}

TEST(FormulonCApiProtection, UpdateMergesIntoExistingRecord) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_protection_t in = FM_STRUCT_INIT(fm_sheet_protection_t);
  in.algorithm_name = "SHA-512";
  in.spin_count = 100000;
  in.allow_format_cells = 1;
  ASSERT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &in), 0);

  fm_sheet_protection_patch_t patch = FM_STRUCT_INIT(fm_sheet_protection_patch_t);
  patch.allow_sort_op = FM_PATCH_OP_SET;
  patch.allow_sort = 1;
  patch.spin_count_op = FM_PATCH_OP_CLEAR;
  patch.algorithm_name_op = FM_PATCH_OP_KEEP;
  ASSERT_EQ(fm_sheet_update_sheet_protection(wb.handle, 0, &patch), 0);

  int32_t present = 0;
  fm_sheet_protection_t out = FM_STRUCT_INIT(fm_sheet_protection_t);
  ASSERT_EQ(fm_sheet_get_sheet_protection(wb.handle, 0, &present, &out), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(out.allow_sort, 1);
  EXPECT_EQ(out.allow_format_cells, 1);
  EXPECT_STREQ(out.algorithm_name, "SHA-512");
  EXPECT_EQ(out.spin_count, 0U);
}

TEST(FormulonCApiProtection, UpdateTakesSelectFlagsAsPlainBooleans) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sheet_protection_t in = FM_STRUCT_INIT(fm_sheet_protection_t);
  ASSERT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &in), 0);

  fm_sheet_protection_patch_t patch = FM_STRUCT_INIT(fm_sheet_protection_patch_t);
  patch.allow_select_locked_cells_op = FM_PATCH_OP_SET;
  patch.allow_select_locked_cells = 0;
  ASSERT_EQ(fm_sheet_update_sheet_protection(wb.handle, 0, &patch), 0);
  int32_t present = 0;
  fm_sheet_protection_t out = FM_STRUCT_INIT(fm_sheet_protection_t);
  ASSERT_EQ(fm_sheet_get_sheet_protection(wb.handle, 0, &present, &out), 0);
  EXPECT_EQ(out.allow_select_locked_cells, FM_TRISTATE_FALSE);
  EXPECT_EQ(out.allow_select_unlocked_cells, FM_TRISTATE_TRUE);

  patch.allow_select_locked_cells = FM_TRISTATE_TRUE;
  EXPECT_EQ(fm_sheet_update_sheet_protection(wb.handle, 0, &patch), FM_STATUS_INVALID_ARGUMENT);
}

TEST(FormulonCApiProtection, UpdateValidatesWholePatchBeforeChanging) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_protection_t in = FM_STRUCT_INIT(fm_sheet_protection_t);
  ASSERT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &in), 0);

  // CLEAR is only valid on strings and the spin count.
  fm_sheet_protection_patch_t clear_flag = FM_STRUCT_INIT(fm_sheet_protection_patch_t);
  clear_flag.allow_sort_op = FM_PATCH_OP_SET;
  clear_flag.allow_sort = 1;
  clear_flag.allow_format_cells_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_sheet_protection(wb.handle, 0, &clear_flag), kInvalid);

  fm_sheet_protection_patch_t null_string = FM_STRUCT_INIT(fm_sheet_protection_patch_t);
  null_string.allow_sort_op = FM_PATCH_OP_SET;
  null_string.allow_sort = 1;
  null_string.hash_value_op = FM_PATCH_OP_SET;
  null_string.hash_value = nullptr;
  EXPECT_EQ(fm_sheet_update_sheet_protection(wb.handle, 0, &null_string), kInvalid);

  fm_sheet_protection_patch_t bad_op = FM_STRUCT_INIT(fm_sheet_protection_patch_t);
  bad_op.allow_sort_op = static_cast<fm_patch_op_t>(3);
  EXPECT_EQ(fm_sheet_update_sheet_protection(wb.handle, 0, &bad_op), kInvalid);

  fm_sheet_protection_patch_t bad_flag = FM_STRUCT_INIT(fm_sheet_protection_patch_t);
  bad_flag.allow_sort_op = FM_PATCH_OP_SET;
  bad_flag.allow_sort = 5;
  EXPECT_EQ(fm_sheet_update_sheet_protection(wb.handle, 0, &bad_flag), kInvalid);

  fm_sheet_protection_t out = FM_STRUCT_INIT(fm_sheet_protection_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_sheet_protection(wb.handle, 0, &present, &out), 0);
  EXPECT_EQ(out.allow_sort, 0);
  EXPECT_EQ(out.allow_format_cells, 0);
}

TEST(FormulonCApiProtection, UpdateRejectsNullAndOutOfRangeSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_protection_patch_t patch = FM_STRUCT_INIT(fm_sheet_protection_patch_t);
  EXPECT_EQ(fm_sheet_update_sheet_protection(nullptr, 0, &patch),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_update_sheet_protection(wb.handle, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_update_sheet_protection(wb.handle, 99, &patch),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kSheetIndexOutOfRange));
}

TEST(FormulonCApiProtection, OutOfRangeSheetReturnsSheetIndexOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  int32_t present = 0;
  fm_sheet_protection_t p = FM_STRUCT_INIT(fm_sheet_protection_t);
  EXPECT_EQ(fm_sheet_get_sheet_protection(wb.handle, 99, &present, &p),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kSheetIndexOutOfRange));
  EXPECT_EQ(fm_sheet_set_sheet_protection(wb.handle, 99, &p),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kSheetIndexOutOfRange));
}

TEST(FormulonCApiProtection, NullArgsReturnBindingNullPointer) {
  int32_t present = 0;
  fm_sheet_protection_t p = FM_STRUCT_INIT(fm_sheet_protection_t);
  EXPECT_EQ(fm_sheet_get_sheet_protection(nullptr, 0, &present, &p),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_set_sheet_protection(nullptr, 0, &p),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_get_sheet_protection(wb.handle, 0, nullptr, &p),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_get_sheet_protection(wb.handle, 0, &present, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiProtection, ProtectedSheetRoundTripsThroughOoxml) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_sheet_protection_t in = FM_STRUCT_INIT(fm_sheet_protection_t);
  in.algorithm_name = "SHA-512";
  in.hash_value = "AAAA==";
  in.salt_value = "BBBB==";
  in.spin_count = 100000;
  in.allow_edit_objects = 1;
  in.allow_format_cells = 1;
  in.allow_select_locked_cells = FM_TRISTATE_FALSE;
  ASSERT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &in), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=1"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);

  int32_t present = 0;
  fm_sheet_protection_t out = FM_STRUCT_INIT(fm_sheet_protection_t);
  ASSERT_EQ(fm_sheet_get_sheet_protection(wb2.handle, 0, &present, &out), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(out.algorithm_name, "SHA-512");
  EXPECT_STREQ(out.hash_value, "AAAA==");
  EXPECT_STREQ(out.salt_value, "BBBB==");
  EXPECT_EQ(out.spin_count, 100000U);
  EXPECT_EQ(out.allow_edit_objects, 1);
  EXPECT_EQ(out.allow_format_cells, 1);
  EXPECT_EQ(out.allow_select_locked_cells, FM_TRISTATE_FALSE);
  EXPECT_EQ(out.allow_select_unlocked_cells, FM_TRISTATE_TRUE);
  // Defaults preserved through round-trip.
  EXPECT_EQ(out.allow_format_rows, 0);
  EXPECT_EQ(out.allow_edit_scenarios, 0);
}

TEST(FormulonCApiProtection, UnprotectedSheetWritesNoProtectionElement) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_protection_t in = FM_STRUCT_INIT(fm_sheet_protection_t);
  in.allow_sort = 1;
  ASSERT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &in), 0);
  ASSERT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, nullptr), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  EXPECT_EQ(IsProtected(wb2.handle), 0);
}

TEST(FormulonCApiProtection, LegacyPasswordHashRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_sheet_protection_t in = FM_STRUCT_INIT(fm_sheet_protection_t);
  in.legacy_password_hash = "CC53";
  ASSERT_EQ(fm_sheet_set_sheet_protection(wb.handle, 0, &in), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  int32_t present = 0;
  fm_sheet_protection_t out = FM_STRUCT_INIT(fm_sheet_protection_t);
  ASSERT_EQ(fm_sheet_get_sheet_protection(wb2.handle, 0, &present, &out), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(out.legacy_password_hash, "CC53");
}
