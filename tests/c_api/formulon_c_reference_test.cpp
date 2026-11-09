//
// Stable C ABI tests for reference text: parsing with INDIRECT's grammar under
// the workbook's profile, A1 formatting, their round trip, and the refusals.

#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "eval/a1_parse.h"
#include "formulon_c_record_arg.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t Code(formulon::FormulonErrorCode c) {
  return static_cast<fm_status_t>(c);
}
constexpr fm_status_t kInvalidArgument = Code(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = Code(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kBindingUtf8 = Code(formulon::FormulonErrorCode::kBindingUtf8EncodingError);

constexpr uint32_t kLastRow = 1048575U;
constexpr uint32_t kLastCol = 16383U;

struct Parsed {
  bool valid = false;
  bool has_sheet = false;
  std::string sheet;
  uint32_t row0 = 0;
  uint32_t col0 = 0;
  uint32_t row1 = 0;
  uint32_t col1 = 0;
};

// Calls the parser with an explicit spec; `r1c1` is an fm_tristate_t value.
fm_status_t ParseSpec(const fm_workbook_t* wb, const char* text, int32_t r1c1, int32_t has_base, uint32_t base_row,
                      uint32_t base_col, fm_parsed_reference_info_t* out) {
  const fm_parse_reference_spec_t spec{sizeof(fm_parse_reference_spec_t), r1c1, has_base, base_row, base_col};
  return fm_workbook_parse_reference(wb, text, &spec, out);
}

// `r1c1` is 0 for A1 and 1 for R1C1 here; it maps onto the tristate.
fm_status_t ParseMode(const fm_workbook_t* wb, const char* text, int32_t r1c1, int32_t has_base, uint32_t base_row,
                      uint32_t base_col, fm_parsed_reference_info_t* out) {
  return ParseSpec(wb, text, r1c1 != 0 ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE, has_base, base_row, base_col, out);
}

// Text that is not a reference fails with the whole text as the parse-error range.
void ExpectNotReference(fm_workbook_t* wb, const char* text, int32_t r1c1 = 0) {
  fm_parsed_reference_info_t out = FM_STRUCT_INIT(fm_parsed_reference_info_t);
  EXPECT_EQ(ParseMode(wb, text, r1c1, 0, 0, 0, &out), FM_STATUS_PARSER_INVALID_REFERENCE) << text;
  EXPECT_EQ(out.sheet_name, nullptr) << text;
  fm_parse_error_info_t error = FM_STRUCT_INIT(fm_parse_error_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_get_last_parse_error(&present, &error), 0) << text;
  ASSERT_EQ(present, 1) << text;
  EXPECT_EQ(error.code, FM_PARSE_ERROR_CODE_INVALID_REFERENCE) << text;
  EXPECT_EQ(error.start, 0U) << text;
  EXPECT_EQ(error.end, static_cast<uint32_t>(std::string(text).size())) << text;
}

Parsed Parse(fm_workbook_t* wb, const char* text, int32_t r1c1 = 0, int32_t has_base = 0, uint32_t base_row = 0,
             uint32_t base_col = 0) {
  fm_parsed_reference_info_t out = FM_STRUCT_INIT(fm_parsed_reference_info_t);
  const fm_status_t rc = ParseMode(wb, text, r1c1, has_base, base_row, base_col, &out);
  EXPECT_EQ(rc, 0) << text << ": " << fm_get_last_error_message();
  Parsed p;
  p.valid = rc == 0;
  p.has_sheet = out.sheet_name != nullptr;
  if (rc != 0) {
    EXPECT_EQ(out.sheet_name, nullptr) << text;
    EXPECT_EQ(out.rect.first_row | out.rect.first_col | out.rect.last_row | out.rect.last_col, 0U) << text;
  }
  if (out.sheet_name != nullptr) {
    p.sheet = out.sheet_name;
  }
  p.row0 = out.rect.first_row;
  p.col0 = out.rect.first_col;
  p.row1 = out.rect.last_row;
  p.col1 = out.rect.last_col;
  return p;
}

void ExpectRange(const Parsed& p, uint32_t row0, uint32_t col0, uint32_t row1, uint32_t col1) {
  ASSERT_TRUE(p.valid);
  EXPECT_EQ(p.row0, row0);
  EXPECT_EQ(p.col0, col0);
  EXPECT_EQ(p.row1, row1);
  EXPECT_EQ(p.col1, col1);
}

std::string Format(fm_workbook_t* wb, const char* sheet, uint32_t row0, uint32_t col0, uint32_t row1, uint32_t col1,
                   int32_t absolute = 0) {
  const char* text = nullptr;
  EXPECT_EQ(
      fm_workbook_format_reference(wb, sheet, record_arg(fm_cell_rect_t{row0, col0, row1, col1}), absolute, &text), 0)
      << fm_get_last_error_message();
  return text != nullptr ? std::string(text) : std::string("<null>");
}

}  // namespace

TEST(FormulonCApiReference, ParsesA1Shapes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  Parsed p = Parse(wb.handle, "A1");
  ExpectRange(p, 0, 0, 0, 0);
  EXPECT_FALSE(p.has_sheet);
  ExpectRange(Parse(wb.handle, "$B$2:C3"), 1, 1, 2, 2);
  ExpectRange(Parse(wb.handle, "C3:B2"), 1, 1, 2, 2);
  ExpectRange(Parse(wb.handle, "B:D"), 0, 1, kLastRow, 3);
  ExpectRange(Parse(wb.handle, "2:11"), 1, 0, 10, kLastCol);
  ExpectRange(Parse(wb.handle, "XFD1048576"), kLastRow, kLastCol, kLastRow, kLastCol);
}

TEST(FormulonCApiReference, ParsesSheetQualifiers) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  struct Row {
    const char* text;
    const char* sheet;
  };
  const Row rows[] = {
      {"Sheet1!A1", "Sheet1"},
      {"'My Sheet'!A1:B2", "My Sheet"},
      {"'O''Brien'!A1", "O'Brien"},
      {"'a!b'!A1", "a!b"},
      {"売上!A1", "売上"},
      {"Sheet.1!A1", "Sheet.1"},
      {"'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'!A1", "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"},
  };
  for (const Row& row : rows) {
    const Parsed p = Parse(wb.handle, row.text);
    ASSERT_TRUE(p.valid) << row.text;
    EXPECT_TRUE(p.has_sheet) << row.text;
    EXPECT_EQ(p.sheet, row.sheet) << row.text;
    EXPECT_EQ(p.row0, 0U) << row.text;
    EXPECT_EQ(p.col0, 0U) << row.text;
  }
  ExpectRange(Parse(wb.handle, "'My Sheet'!A1:B2"), 0, 0, 1, 1);
}

TEST(FormulonCApiReference, NonReferencesFailAsInvalidReference) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* rows[] = {
      "XFE1",
      "A1048577",
      "A0",
      "Sheet1:Sheet2!A1",
      "[1]Sheet1!A1",
      "MyName",
      "Table1[Col]",
      "A1,B2",
      "",
      "A1 B2",
      " A1",
      "A1 ",
      // 32 UTF-16 units, one past Excel's sheet-name limit.
      "'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'!A1",
  };
  for (const char* text : rows) {
    ExpectNotReference(wb.handle, text);
  }
}

TEST(FormulonCApiReference, ParsesR1C1UnderTheWorkbookProfile) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  ExpectRange(Parse(wb.handle, "R2C3", 1), 1, 2, 1, 2);
  ExpectRange(Parse(wb.handle, "R[-1]C", 1, 1, 5, 5), 4, 5, 4, 5);
  // A relative axis wraps at the grid edge, as Excel's does.
  ExpectRange(Parse(wb.handle, "R[-1]C", 1, 1, 0, 0), kLastRow, 0, kLastRow, 0);
  ExpectNotReference(wb.handle, "R[-1]C", 1);

  WorkbookGuard de;
  ASSERT_EQ(fm_workbook_create_default(&de.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(de.handle, "mac-365-de_DE"), 0);
  ExpectRange(Parse(de.handle, "Z2S3", 1), 1, 2, 1, 2);
}

TEST(FormulonCApiReference, SheetNameOutlivesTheNextParserCall) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_parsed_reference_info_t out = FM_STRUCT_INIT(fm_parsed_reference_info_t);
  ASSERT_EQ(fm_workbook_parse_reference(wb.handle, "'My Sheet'!A1", nullptr, &out), 0);
  ASSERT_NE(out.sheet_name, nullptr);
  // The parser's own sheet view lives in thread-local scratch the next parse reuses.
  const auto other = formulon::eval::refs_internal::parse_a1_ref("'Other one'!B2");
  ASSERT_TRUE(other.valid);
  EXPECT_STREQ(out.sheet_name, "My Sheet");
}

TEST(FormulonCApiReference, ParseRejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_parsed_reference_info_t out = FM_STRUCT_INIT(fm_parsed_reference_info_t);
  EXPECT_EQ(fm_workbook_parse_reference(nullptr, "A1", nullptr, &out), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_parse_reference(wb.handle, nullptr, nullptr, &out), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_parse_reference(wb.handle, "A1", nullptr, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_parse_reference(wb.handle, "A\xff", nullptr, &out), kBindingUtf8);
  EXPECT_EQ(ParseSpec(wb.handle, "A1", 3, 0, 0, 0, &out), kInvalidArgument);
  EXPECT_EQ(ParseSpec(wb.handle, "A1", -1, 0, 0, 0, &out), kInvalidArgument);
  EXPECT_EQ(ParseSpec(wb.handle, "A1", FM_TRISTATE_DEFAULT, 2, 0, 0, &out), kInvalidArgument);
  // The base is checked in A1 mode too.
  EXPECT_EQ(ParseSpec(wb.handle, "A1", FM_TRISTATE_DEFAULT, 1, kLastRow + 1U, 0, &out), kInvalidArgument);
  EXPECT_EQ(ParseSpec(wb.handle, "RC", FM_TRISTATE_TRUE, 1, 0, kLastCol + 1U, &out), kInvalidArgument);
}

TEST(FormulonCApiReference, NullSpecParsesPlainA1WithoutBase) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_parsed_reference_info_t out = FM_STRUCT_INIT(fm_parsed_reference_info_t);
  ASSERT_EQ(fm_workbook_parse_reference(wb.handle, "B2:C3", nullptr, &out), FM_STATUS_OK)
      << fm_get_last_error_message();
  EXPECT_EQ(out.sheet_name, nullptr);
  EXPECT_EQ(out.rect.first_row, 1U);
  EXPECT_EQ(out.rect.first_col, 1U);
  EXPECT_EQ(out.rect.last_row, 2U);
  EXPECT_EQ(out.rect.last_col, 2U);
}

TEST(FormulonCApiReference, DefaultAndFalseTristateParseA1AndTrueParsesR1C1) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (const int32_t a1_mode : {static_cast<int32_t>(FM_TRISTATE_DEFAULT), static_cast<int32_t>(FM_TRISTATE_FALSE)}) {
    fm_parsed_reference_info_t out = FM_STRUCT_INIT(fm_parsed_reference_info_t);
    ASSERT_EQ(ParseSpec(wb.handle, "B2", a1_mode, 0, 0, 0, &out), FM_STATUS_OK) << fm_get_last_error_message();
    EXPECT_EQ(out.rect.first_row, 1U);
    EXPECT_EQ(out.rect.first_col, 1U);
  }
  fm_parsed_reference_info_t r1c1 = FM_STRUCT_INIT(fm_parsed_reference_info_t);
  ASSERT_EQ(ParseSpec(wb.handle, "R2C3", FM_TRISTATE_TRUE, 0, 0, 0, &r1c1), FM_STATUS_OK)
      << fm_get_last_error_message();
  EXPECT_EQ(r1c1.rect.first_row, 1U);
  EXPECT_EQ(r1c1.rect.first_col, 2U);
}

TEST(FormulonCApiReference, R1c1RelativeAxisWithoutBaseIsNotAReference) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_parsed_reference_info_t out = FM_STRUCT_INIT(fm_parsed_reference_info_t);
  EXPECT_EQ(ParseSpec(wb.handle, "R[1]C[1]", FM_TRISTATE_TRUE, 0, 0, 0, &out), FM_STATUS_PARSER_INVALID_REFERENCE);
  ASSERT_EQ(ParseSpec(wb.handle, "R[1]C[1]", FM_TRISTATE_TRUE, 1, 4, 4, &out), FM_STATUS_OK)
      << fm_get_last_error_message();
  EXPECT_EQ(out.rect.first_row, 5U);
  EXPECT_EQ(out.rect.first_col, 5U);
}

TEST(FormulonCApiReference, FormatsA1) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  EXPECT_EQ(Format(wb.handle, nullptr, 0, 0, 0, 0), "A1");
  EXPECT_EQ(Format(wb.handle, "", 0, 0, 0, 0), "A1");
  EXPECT_EQ(Format(wb.handle, nullptr, 0, 1, kLastRow, 3), "B:D");
  EXPECT_EQ(Format(wb.handle, nullptr, 1, 0, 10, kLastCol), "2:11");
  EXPECT_EQ(Format(wb.handle, "My Sheet", 0, 0, 0, 0), "'My Sheet'!A1");
  EXPECT_EQ(Format(wb.handle, "A1", 0, 0, 0, 0), "'A1'!A1");
  EXPECT_EQ(Format(wb.handle, "TRUE", 0, 0, 0, 0), "'TRUE'!A1");
  EXPECT_EQ(Format(wb.handle, "1a", 0, 0, 0, 0), "'1a'!A1");
  EXPECT_EQ(Format(wb.handle, "売上", 0, 0, 0, 0), "売上!A1");
  EXPECT_EQ(Format(wb.handle, "Sheet1", 0, 0, 0, 0), "Sheet1!A1");
  EXPECT_EQ(Format(wb.handle, nullptr, 0, 0, 1, 1, 1), "$A$1:$B$2");
}

TEST(FormulonCApiReference, FormatRejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* text = nullptr;
  EXPECT_EQ(fm_workbook_format_reference(nullptr, nullptr, record_arg(fm_cell_rect_t{0, 0, 0, 0}), 0, &text),
            kBindingNullPointer);
  EXPECT_EQ(fm_workbook_format_reference(wb.handle, nullptr, record_arg(fm_cell_rect_t{0, 0, 0, 0}), 0, nullptr),
            kBindingNullPointer);
  EXPECT_EQ(fm_workbook_format_reference(wb.handle, "S\xff", record_arg(fm_cell_rect_t{0, 0, 0, 0}), 0, &text),
            kBindingUtf8);
  EXPECT_EQ(
      fm_workbook_format_reference(wb.handle, nullptr, record_arg(fm_cell_rect_t{0, 0, kLastRow + 1U, 0}), 0, &text),
      kInvalidArgument);
  EXPECT_EQ(
      fm_workbook_format_reference(wb.handle, nullptr, record_arg(fm_cell_rect_t{0, 0, 0, kLastCol + 1U}), 0, &text),
      kInvalidArgument);
  EXPECT_EQ(fm_workbook_format_reference(wb.handle, nullptr, record_arg(fm_cell_rect_t{2, 0, 1, 0}), 0, &text),
            kInvalidArgument);
  EXPECT_EQ(fm_workbook_format_reference(wb.handle, nullptr, record_arg(fm_cell_rect_t{0, 2, 0, 1}), 0, &text),
            kInvalidArgument);
  EXPECT_EQ(fm_workbook_format_reference(wb.handle, nullptr, record_arg(fm_cell_rect_t{0, 0, 0, 0}), 2, &text),
            kInvalidArgument);
}

TEST(FormulonCApiReference, FormatThenParseRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* sheets[] = {"Sheet1", "My Sheet", "O'Brien", "売上", "Sheet.1", "1a", "A1", "TRUE"};
  struct Rect {
    uint32_t row0, col0, row1, col1;
  };
  const Rect rects[] = {{0, 0, 0, 0}, {1, 1, 2, 2}, {0, 1, kLastRow, 3}, {1, 0, 10, kLastCol}};
  for (const char* sheet : sheets) {
    for (const Rect& r : rects) {
      for (const int32_t absolute : {0, 1}) {
        const std::string text = Format(wb.handle, sheet, r.row0, r.col0, r.row1, r.col1, absolute);
        const Parsed p = Parse(wb.handle, text.c_str());
        ASSERT_TRUE(p.valid) << text;
        EXPECT_EQ(p.sheet, sheet) << text;
        EXPECT_EQ(p.row0, r.row0) << text;
        EXPECT_EQ(p.col0, r.col0) << text;
        EXPECT_EQ(p.row1, r.row1) << text;
        EXPECT_EQ(p.col1, r.col1) << text;
      }
    }
  }
}
