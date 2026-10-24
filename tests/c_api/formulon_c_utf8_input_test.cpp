// Stable C ABI text input: malformed UTF-8 is refused with
// kBindingUtf8EncodingError before anything is stored, and well-formed
// multibyte text passes.

#include <cstddef>
#include <cstdint>
#include <string>

#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kUtf8Error = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingUtf8EncodingError);

// U+D800 encoded as if it were a scalar: what an embind string carries for a
// lone UTF-16 surrogate.
constexpr const char* kLoneSurrogate = "a\xED\xA0\x80";
// The first two bytes of the three-byte encoding of U+3042.
constexpr const char* kTruncated = "\xE3\x81";
// U+3042 U+1F600: a three-byte and a four-byte scalar.
constexpr const char* kValidMultibyte = "\xE3\x81\x82\xF0\x9F\x98\x80";

const char* const kMalformed[] = {kLoneSurrogate, kTruncated};

std::size_t CellCount(const WorkbookGuard& wb) {
  std::size_t cells = 99U;
  EXPECT_EQ(fm_workbook_cell_count(wb.handle, 0, &cells), 0) << fm_last_error_message();
  return cells;
}

std::string SheetName(const WorkbookGuard& wb, std::size_t index) {
  const char* name = nullptr;
  EXPECT_EQ(fm_workbook_sheet_name(wb.handle, index, &name), 0) << fm_last_error_message();
  return name != nullptr ? name : "";
}

std::uint32_t CommentCount(const WorkbookGuard& wb) {
  std::uint32_t count = 99U;
  EXPECT_EQ(fm_sheet_get_comment_count(wb.handle, 0, &count), 0) << fm_last_error_message();
  return count;
}

}  // namespace

TEST(FormulonCApiUtf8Input, MalformedTextIsRefusedWithTheWorkbookUnchanged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const std::string sheet_name = SheetName(wb, 0);
  for (const char* bad : kMalformed) {
    EXPECT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, bad), kUtf8Error);
    EXPECT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 1, (std::string("=\"") + bad + "\"").c_str()), kUtf8Error);
    EXPECT_EQ(fm_workbook_rename_sheet(wb.handle, 0, bad), kUtf8Error);
    EXPECT_EQ(fm_workbook_add_sheet(wb.handle, bad), kUtf8Error);
    EXPECT_EQ(fm_sheet_set_comment(wb.handle, 0, 0, 2, "author", bad), kUtf8Error);
    EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, bad, "=1"), kUtf8Error);
    EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, "Name", (std::string("=\"") + bad + "\"").c_str()), kUtf8Error);
    // The diagnostic never echoes the rejected bytes.
    EXPECT_STREQ(fm_last_error_message(), "fm_workbook_set_defined_name: text is not valid UTF-8");
  }
  EXPECT_EQ(CellCount(wb), 0U);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 1U);
  EXPECT_EQ(SheetName(wb, 0), sheet_name);
  EXPECT_EQ(CommentCount(wb), 0U);
  EXPECT_EQ(fm_workbook_defined_name_count(wb.handle), 0U);
}

TEST(FormulonCApiUtf8Input, MalformedTextIsRefusedByReadOnlyEntryPoints) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  for (const char* bad : kMalformed) {
    std::int32_t order = 99;
    EXPECT_EQ(fm_text_compare(bad, "a", "win-365-en_US", FM_COLLATION_COMPARE, &order), kUtf8Error);
    EXPECT_EQ(fm_text_compare("a", bad, "win-365-en_US", FM_COLLATION_COMPARE, &order), kUtf8Error);
    EXPECT_EQ(order, 99);
    fm_value_t value{};
    EXPECT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 0, (std::string("=LEN(\"") + bad + "\")").c_str(), &value),
              kUtf8Error);
    fm_value_t text{};
    text.kind = FM_VAL_TEXT;
    text.u.text = bad;
    const char* shown = nullptr;
    std::int32_t status = 0;
    EXPECT_EQ(fm_workbook_format_value(wb.handle, &text, "@", &shown, &status), kUtf8Error);
  }
}

TEST(FormulonCApiUtf8Input, WellFormedMultibyteTextPasses) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, kValidMultibyte), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_rename_sheet(wb.handle, 0, kValidMultibyte), 0) << fm_last_error_message();
  ASSERT_EQ(fm_sheet_set_comment(wb.handle, 0, 0, 1, "author", kValidMultibyte), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 2, (std::string("=\"") + kValidMultibyte + "\"").c_str()), 0)
      << fm_last_error_message();
  std::int32_t order = 99;
  ASSERT_EQ(fm_text_compare(kValidMultibyte, kValidMultibyte, "win-365-en_US", FM_COLLATION_COMPARE, &order), 0)
      << fm_last_error_message();
  EXPECT_EQ(order, 0);

  fm_value_t value{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &value), 0);
  ASSERT_EQ(value.kind, FM_VAL_TEXT);
  EXPECT_STREQ(value.u.text, kValidMultibyte);
  EXPECT_EQ(SheetName(wb, 0), kValidMultibyte);
  EXPECT_EQ(CommentCount(wb), 1U);
}
