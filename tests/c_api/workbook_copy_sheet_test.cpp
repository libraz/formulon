//
// Stable C ABI tests for fm_workbook_copy_sheet: placement, naming and every
// refusal the boundary contract names.

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_record_arg.h"
#include "gtest/gtest.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

using formulon::FormulonErrorCode;

constexpr fm_status_t Code(FormulonErrorCode code) {
  return static_cast<fm_status_t>(code);
}

constexpr std::uint32_t kSentinel = 0xDEADBEEFU;

std::string SheetName(const fm_workbook_t* wb, std::size_t index) {
  const char* name = nullptr;
  EXPECT_EQ(fm_workbook_get_sheet_name(wb, index, &name), 0);
  return name == nullptr ? std::string() : std::string(name);
}

std::size_t SheetCount(const fm_workbook_t* wb) {
  return count_of(fm_workbook_get_sheet_count, wb);
}

/// Sheet1 holding A1 = 5 and B1 = A1 * 2.
void MakeSingleSheet(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 5.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "=A1*2"), 0);
}

/// Copies sheet 0 before `insert_before` and expects `status`, leaving the
/// workbook untouched on failure.
void ExpectCopyStatus(fm_workbook_t* wb, std::uint32_t source, std::uint32_t insert_before, const char* name,
                      fm_status_t status) {
  const std::size_t before = SheetCount(wb);
  std::size_t out = kSentinel;
  EXPECT_EQ(fm_workbook_copy_sheet(wb, source, insert_before, name, &out), status) << fm_get_last_error_message();
  if (status != 0) {
    EXPECT_EQ(SheetCount(wb), before);
    EXPECT_EQ(out, kSentinel);
  }
}

}  // namespace

TEST(CApiCopySheet, SingleSheetAutomaticName) {
  WorkbookGuard wb;
  MakeSingleSheet(wb);
  std::size_t out = kSentinel;
  ASSERT_EQ(fm_workbook_copy_sheet(wb.handle, 0, 1, nullptr, &out), 0) << fm_get_last_error_message();
  EXPECT_EQ(out, 1U);
  ASSERT_EQ(SheetCount(wb.handle), 2U);
  EXPECT_EQ(SheetName(wb.handle, 1), "Sheet1 (2)");

  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 1, 0, 1, &value), 0);
  EXPECT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(value.u.number, 10.0);
}

TEST(CApiCopySheet, ExplicitNameAndInsertBeforeFirst) {
  WorkbookGuard wb;
  MakeSingleSheet(wb);
  std::size_t out = kSentinel;
  ASSERT_EQ(fm_workbook_copy_sheet(wb.handle, 0, 0, "Archive", &out), 0) << fm_get_last_error_message();
  EXPECT_EQ(out, 0U);
  EXPECT_EQ(SheetName(wb.handle, 0), "Archive");
  EXPECT_EQ(SheetName(wb.handle, 1), "Sheet1");
}

TEST(CApiCopySheet, AppendAtSheetCountAfterReindex) {
  WorkbookGuard wb;
  MakeSingleSheet(wb);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Other", ignored_index()), 0);
  std::size_t out = kSentinel;
  ASSERT_EQ(fm_workbook_copy_sheet(wb.handle, 0, 2, "Tail", &out), 0) << fm_get_last_error_message();
  EXPECT_EQ(out, 2U);
  ASSERT_EQ(SheetCount(wb.handle), 3U);
  EXPECT_EQ(SheetName(wb.handle, 1), "Other");
  EXPECT_EQ(SheetName(wb.handle, 2), "Tail");
}

TEST(CApiCopySheet, IndexOutOfRange) {
  WorkbookGuard wb;
  MakeSingleSheet(wb);
  const fm_status_t oor = Code(FormulonErrorCode::kSheetIndexOutOfRange);
  ExpectCopyStatus(wb.handle, 1, 1, nullptr, oor);
  ExpectCopyStatus(wb.handle, 0, 2, nullptr, oor);
  ExpectCopyStatus(wb.handle, 0x80000000U, 0, nullptr, oor);
  ExpectCopyStatus(wb.handle, 0xFFFFFFFFU, 0, nullptr, oor);
  ExpectCopyStatus(wb.handle, 0, 0x80000000U, nullptr, oor);
  ExpectCopyStatus(wb.handle, 0, 0xFFFFFFFFU, nullptr, oor);
}

TEST(CApiCopySheet, SheetNameValidation) {
  WorkbookGuard wb;
  MakeSingleSheet(wb);
  const fm_status_t invalid = Code(FormulonErrorCode::kInvalidSheetName);
  ExpectCopyStatus(wb.handle, 0, 1, "", invalid);
  ExpectCopyStatus(wb.handle, 0, 1, "SHEET1", invalid);
  ExpectCopyStatus(wb.handle, 0, 1, "a:b", invalid);
  ExpectCopyStatus(wb.handle, 0, 1, std::string(32U, 'x').c_str(), invalid);
  ExpectCopyStatus(wb.handle, 0, 1, std::string(31U, 'x').c_str(), 0);
}

TEST(CApiCopySheet, NameLengthCountsUtf16Units) {
  WorkbookGuard wb;
  MakeSingleSheet(wb);
  // U+1F600 is two UTF-16 units: 15 of them plus one letter is 31 units, 16 is 32.
  std::string fifteen;
  for (int i = 0; i < 15; ++i) {
    fifteen += "\xF0\x9F\x98\x80";
  }
  ExpectCopyStatus(wb.handle, 0, 1, (fifteen + "a").c_str(), 0);
  ExpectCopyStatus(wb.handle, 0, 1, (fifteen + "\xF0\x9F\x98\x80").c_str(), Code(FormulonErrorCode::kInvalidSheetName));
}

TEST(CApiCopySheet, MalformedUtf8NameIsRejected) {
  WorkbookGuard wb;
  MakeSingleSheet(wb);
  ExpectCopyStatus(wb.handle, 0, 1, "\xC3\x28", Code(FormulonErrorCode::kBindingUtf8EncodingError));
}

TEST(CApiCopySheet, NullArguments) {
  WorkbookGuard wb;
  MakeSingleSheet(wb);
  const fm_status_t null_pointer = Code(FormulonErrorCode::kBindingNullPointer);
  std::size_t out = kSentinel;
  EXPECT_EQ(fm_workbook_copy_sheet(nullptr, 0, 1, nullptr, &out), null_pointer);
  EXPECT_EQ(fm_workbook_copy_sheet(wb.handle, 0, 1, nullptr, nullptr), null_pointer);
  EXPECT_EQ(SheetCount(wb.handle), 1U);
}

TEST(CApiCopySheet, SheetCountLimit) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  formulon::Workbook& model = wb.handle->workbook();
  std::string name;
  while (model.sheet_count() < formulon::Workbook::kMaxSheets) {
    name.assign("S");
    name.append(std::to_string(model.sheet_count()));
    model.add_sheet(name);
  }
  ExpectCopyStatus(wb.handle, 0, 1, nullptr, Code(FormulonErrorCode::kSheetCountLimitExceeded));
}

TEST(CApiCopySheet, VeryHiddenSourceRefused) {
  WorkbookGuard wb;
  MakeSingleSheet(wb);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Visible", ignored_index()), 0);
  ASSERT_EQ(fm_sheet_set_visibility(wb.handle, 0, FM_SHEET_VISIBILITY_VERY_HIDDEN), 0) << fm_get_last_error_message();
  ExpectCopyStatus(wb.handle, 0, 2, nullptr, Code(FormulonErrorCode::kPreconditionFailed));
}

TEST(CApiCopySheet, SlicerSourceRefusedWithoutChange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Product"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 1, "Amount"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 1, 0, "Widget"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 10.0), 0);
  const char* columns[] = {"Product", "Amount"};
  std::size_t table = 99U;
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 1, 1}), "Sales", columns, 2, "TableStyleMedium2", 1,
                     0, &table),
            0);
  fm_slicer_t spec = FM_STRUCT_INIT(fm_slicer_t);
  spec.fields = FM_SLICER_FIELD_ANCHOR;
  spec.from_row = 4;
  spec.from_col = 3;
  spec.to_row = 14;
  spec.to_col = 5;
  std::uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, &id), 0) << fm_get_last_error_message();
  ExpectCopyStatus(wb.handle, 0, 1, nullptr, Code(FormulonErrorCode::kNotImplemented));
}
