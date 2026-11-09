// Failure text of the C ABI: messages are English sentences that name no C
// function and no C++ scope, and contexts stay structured `key=value` pairs.

#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_out_index.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

std::string Message() {
  return fm_get_last_error_message();
}

std::string Context() {
  return fm_get_last_error_context();
}

void ExpectNoInternalNames(const std::string& text) {
  EXPECT_EQ(text.find("fm_"), std::string::npos) << text;
  EXPECT_EQ(text.find("::"), std::string::npos) << text;
}

}  // namespace

TEST(FormulonCApiErrorText, SheetIndexOutOfRangeNamesTheIndexAndTheSheetCount) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cell_count(wb.handle, 5, &count), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(Message(), "sheet index 5 is out of range (1 sheet)");
  EXPECT_EQ(Context(), "sheet_index=5 sheet_count=1");
  ExpectNoInternalNames(Message());
  ExpectNoInternalNames(Context());

  size_t added = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second", append_position(wb.handle), &added), 0);
  ASSERT_EQ(fm_sheet_get_cell_count(wb.handle, 7, &count), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(Message(), "sheet index 7 is out of range (2 sheets)");
  EXPECT_EQ(Context(), "sheet_index=7 sheet_count=2");

  const char* name = nullptr;
  ASSERT_EQ(fm_workbook_get_sheet_name(wb.handle, 2, &name), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(Message(), "sheet index 2 is out of range (2 sheets)");
}

TEST(FormulonCApiErrorText, FontIndexOutOfRangeNamesTheFieldTheIndexAndTheCount) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  size_t font_count = 0;
  ASSERT_EQ(fm_workbook_get_font_count(wb.handle, &font_count), 0);
  fm_font_t record{};
  ASSERT_EQ(fm_workbook_get_font_at(wb.handle, 999, &record), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(Message(), "font_index 999 is out of range (" + std::to_string(font_count) + " items)");
  EXPECT_EQ(Context(), "font_index=999 count=" + std::to_string(font_count));
  ExpectNoInternalNames(Message());
  ExpectNoInternalNames(Context());
}

TEST(FormulonCApiErrorText, NullArgumentMessagesNameTheArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t count = 0;
  EXPECT_EQ(fm_workbook_get_sheet_count(nullptr, &count), FM_STATUS_BINDING_NULL_POINTER);
  ExpectNoInternalNames(Message());
  EXPECT_NE(Message().find("NULL"), std::string::npos) << Message();

  fm_workbook_t* loaded = nullptr;
  const uint8_t byte = 0;
  EXPECT_EQ(fm_workbook_load(nullptr, 4, &loaded), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(Message(), "bytes is NULL");
  EXPECT_EQ(fm_workbook_load(&byte, 1, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(Message(), "out is NULL");
}

TEST(FormulonCApiErrorText, InvalidArgumentMessagesAreSentences) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_error(wb.handle, 0, 0, 0, 9999), FM_STATUS_INVALID_ARGUMENT);
  ExpectNoInternalNames(Message());
  ExpectNoInternalNames(Context());
  EXPECT_FALSE(Message().empty());
}

TEST(FormulonCApiErrorText, CorruptArchiveReportsTheZipCodeWithoutInternalNames) {
  std::vector<uint8_t> bytes = {'P', 'K', 0x03, 0x04};
  bytes.insert(bytes.end(), 12, 0);
  ASSERT_EQ(bytes.size(), 16U);
  fm_workbook_t* wb = nullptr;
  EXPECT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), FM_STATUS_IO_ZIP_CORRUPT);
  EXPECT_EQ(wb, nullptr);
  ExpectNoInternalNames(Message());
  ExpectNoInternalNames(Context());
  EXPECT_FALSE(Message().empty());
}

TEST(FormulonCApiErrorText, MissingNaturalKeyRemoveReportsNotFoundWithoutInternalNames) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_remove_defined_name(wb.handle, "nope", -1), FM_STATUS_NOT_FOUND);
  ExpectNoInternalNames(Message());
  ExpectNoInternalNames(Context());
  EXPECT_FALSE(Message().empty());
  EXPECT_NE(Context().find("name=nope"), std::string::npos) << Context();
}

TEST(FormulonCApiErrorText, MissingSheetNameLeavesNoErrorText) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_remove_defined_name(wb.handle, "nope", -1), FM_STATUS_NOT_FOUND);
  size_t index = 99;
  int32_t present = -1;
  ASSERT_EQ(fm_workbook_get_sheet_index(wb.handle, "No such sheet", &present, &index), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(Message(), "");
  EXPECT_EQ(Context(), "");
}
