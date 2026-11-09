// Stable C ABI text input: malformed UTF-8 is refused with
// kBindingUtf8EncodingError before anything is stored, and well-formed
// multibyte text passes.

#include <cstddef>
#include <cstdint>
#include <string>

#include "formulon_c_evaluate_top_left.h"
#include "formulon_c_out_index.h"
#include "formulon_c_setters.h"
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
  EXPECT_EQ(fm_sheet_get_cell_count(wb.handle, 0, &cells), 0) << fm_get_last_error_message();
  return cells;
}

std::string SheetName(const WorkbookGuard& wb, std::size_t index) {
  const char* name = nullptr;
  EXPECT_EQ(fm_workbook_get_sheet_name(wb.handle, index, &name), 0) << fm_get_last_error_message();
  return name != nullptr ? name : "";
}

std::uint32_t CommentCount(const WorkbookGuard& wb) {
  std::size_t count = 99U;
  EXPECT_EQ(fm_sheet_get_comment_count(wb.handle, 0, &count), 0) << fm_get_last_error_message();
  return static_cast<std::uint32_t>(count);
}

}  // namespace

TEST(FormulonCApiUtf8Input, MalformedTextIsRefusedWithTheWorkbookUnchanged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::string sheet_name = SheetName(wb, 0);
  for (const char* bad : kMalformed) {
    EXPECT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, bad), kUtf8Error);
    EXPECT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, (std::string("=\"") + bad + "\"").c_str()), kUtf8Error);
    EXPECT_EQ(fm_workbook_rename_sheet(wb.handle, 0, bad), kUtf8Error);
    EXPECT_EQ(fm_workbook_add_sheet(wb.handle, bad, ignored_index()), kUtf8Error);
    EXPECT_EQ(fm_sheet_set_comment(wb.handle, 0, 0, 2, "author", bad), kUtf8Error);
    EXPECT_EQ(SetDefinedName(wb.handle, bad, "=1"), kUtf8Error);
    EXPECT_EQ(SetDefinedName(wb.handle, "Name", (std::string("=\"") + bad + "\"").c_str()), kUtf8Error);
    // The diagnostic never echoes the rejected bytes.
    EXPECT_STREQ(fm_get_last_error_message(), "text is not valid UTF-8");
  }
  EXPECT_EQ(CellCount(wb), 0U);
  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), 1U);
  EXPECT_EQ(SheetName(wb, 0), sheet_name);
  EXPECT_EQ(CommentCount(wb), 0U);
  EXPECT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 0U);
}

TEST(FormulonCApiUtf8Input, MalformedTextIsRefusedByReadOnlyEntryPoints) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (const char* bad : kMalformed) {
    std::int32_t order = 99;
    EXPECT_EQ(fm_compare_text(bad, "a", FM_COLLATION_STRENGTH_COMPARE, "win-365-en_US", &order), kUtf8Error);
    EXPECT_EQ(fm_compare_text("a", bad, FM_COLLATION_STRENGTH_COMPARE, "win-365-en_US", &order), kUtf8Error);
    EXPECT_EQ(order, 99);
    fm_value_t value{};
    EXPECT_EQ(evaluate_top_left(wb.handle, 0, 0, 0, (std::string("=LEN(\"") + bad + "\")").c_str(), &value),
              kUtf8Error);
    fm_value_t text{};
    text.kind = FM_VALUE_KIND_TEXT;
    text.u.text = bad;
    const char* shown = nullptr;
    std::int32_t status = 0;
    EXPECT_EQ(fm_workbook_format_value(wb.handle, &text, "@", &shown, &status), kUtf8Error);
  }
}

TEST(FormulonCApiUtf8Input, KeyedGettersRefuseAMalformedKey) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (const char* bad : {kLoneSurrogate, kTruncated, "\xff"}) {
    std::int32_t present = 99;
    fm_defined_name_t name = FM_STRUCT_INIT(fm_defined_name_t);
    EXPECT_EQ(fm_workbook_get_defined_name(wb.handle, bad, -1, &present, &name), kUtf8Error);
    EXPECT_EQ(fm_workbook_has_passthrough(wb.handle, bad, &present), kUtf8Error);
    fm_threaded_comment_t comment = FM_STRUCT_INIT(fm_threaded_comment_t);
    EXPECT_EQ(fm_sheet_get_threaded_comment(wb.handle, 0, bad, &present, &comment), kUtf8Error);
    fm_person_t person = FM_STRUCT_INIT(fm_person_t);
    EXPECT_EQ(fm_workbook_get_person(wb.handle, bad, &present, &person), kUtf8Error);
    fm_table_style_t table_style = FM_STRUCT_INIT(fm_table_style_t);
    EXPECT_EQ(fm_workbook_get_table_style(wb.handle, bad, &present, &table_style), kUtf8Error);
    fm_cell_style_t cell_style{};
    EXPECT_EQ(fm_workbook_get_cell_style(wb.handle, bad, &present, &cell_style), kUtf8Error);
    const char* typeface = nullptr;
    EXPECT_EQ(fm_workbook_get_theme_script_font(wb.handle, FM_FONT_SCHEME_MINOR, bad, &present, &typeface), kUtf8Error);
    EXPECT_EQ(present, 99);
  }
}

// A drawing rename checks its text before it looks up the id, so even an id
// no object holds reports the encoding.
TEST(FormulonCApiUtf8Input, DrawingRenameRefusesAMalformedName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (const char* bad : {kLoneSurrogate, kTruncated, "\xff"}) {
    EXPECT_EQ(fm_sheet_set_drawing_name(wb.handle, 0, 42U, bad), kUtf8Error);
  }
}

TEST(FormulonCApiUtf8Input, FontAndStyleRecordsRefuseAMalformedName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::size_t fonts = count_of(fm_workbook_get_font_count, wb.handle);
  const std::size_t dxfs = count_of(fm_workbook_get_dxf_count, wb.handle);
  for (const char* bad : kMalformed) {
    fm_font_t font{};
    font.name = bad;
    font.size_pt = 11.0;
    std::size_t font_index = 99U;
    EXPECT_EQ(fm_workbook_add_font(wb.handle, &font, &font_index), kUtf8Error);
    EXPECT_EQ(font_index, 99U);
    EXPECT_EQ(fm_workbook_set_font(wb.handle, 0, &font), kUtf8Error);
    fm_dxf_t dxf{};
    dxf.font_engaged = 1;
    dxf.font = font;
    std::uint32_t dxf_index = 99U;
    EXPECT_EQ(fm_workbook_add_dxf(wb.handle, &dxf, &dxf_index), kUtf8Error);
    EXPECT_EQ(dxf_index, 99U);
  }
  EXPECT_EQ(count_of(fm_workbook_get_font_count, wb.handle), fonts);
  EXPECT_EQ(count_of(fm_workbook_get_dxf_count, wb.handle), dxfs);
}

TEST(FormulonCApiUtf8Input, WellFormedMultibyteTextPasses) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, kValidMultibyte), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_rename_sheet(wb.handle, 0, kValidMultibyte), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_comment(wb.handle, 0, 0, 1, "author", kValidMultibyte), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 2, (std::string("=\"") + kValidMultibyte + "\"").c_str()), 0)
      << fm_get_last_error_message();
  std::int32_t order = 99;
  ASSERT_EQ(fm_compare_text(kValidMultibyte, kValidMultibyte, FM_COLLATION_STRENGTH_COMPARE, "win-365-en_US", &order),
            0)
      << fm_get_last_error_message();
  EXPECT_EQ(order, 0);

  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 0, &value), 0);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_TEXT);
  EXPECT_STREQ(value.u.text, kValidMultibyte);
  EXPECT_EQ(SheetName(wb, 0), kValidMultibyte);
  EXPECT_EQ(CommentCount(wb), 1U);
}
