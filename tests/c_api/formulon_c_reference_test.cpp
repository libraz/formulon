//
// Stable C ABI tests for reference text: parsing with INDIRECT's grammar under
// the workbook's profile, A1 formatting, their round trip, and the refusals.

#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "eval/a1_parse.h"
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

Parsed Parse(fm_workbook_t* wb, const char* text, int32_t r1c1 = 0, int32_t has_base = 0, uint32_t base_row = 0,
             uint32_t base_col = 0) {
  fm_parsed_reference_t out{};
  EXPECT_EQ(fm_workbook_parse_reference(wb, text, r1c1, has_base, base_row, base_col, &out), 0)
      << text << ": " << fm_get_last_error_message();
  Parsed p;
  p.valid = out.valid == 1;
  p.has_sheet = out.has_sheet == 1;
  if (out.valid == 0) {
    EXPECT_EQ(out.has_sheet, 0) << text;
    EXPECT_EQ(out.sheet_name, nullptr) << text;
    EXPECT_EQ(out.row0 | out.col0 | out.row1 | out.col1, 0U) << text;
  }
  if (out.has_sheet == 1) {
    EXPECT_NE(out.sheet_name, nullptr) << text;
    if (out.sheet_name != nullptr) {
      p.sheet = out.sheet_name;
    }
  } else {
    EXPECT_EQ(out.sheet_name, nullptr) << text;
  }
  p.row0 = out.row0;
  p.col0 = out.col0;
  p.row1 = out.row1;
  p.col1 = out.col1;
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
  EXPECT_EQ(fm_workbook_format_reference(wb, sheet, row0, col0, row1, col1, absolute, &text), 0)
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

TEST(FormulonCApiReference, NonReferencesReadInvalid) {
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
    EXPECT_FALSE(Parse(wb.handle, text).valid) << text;
  }
}

TEST(FormulonCApiReference, ParsesR1C1UnderTheWorkbookProfile) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  ExpectRange(Parse(wb.handle, "R2C3", 1), 1, 2, 1, 2);
  ExpectRange(Parse(wb.handle, "R[-1]C", 1, 1, 5, 5), 4, 5, 4, 5);
  // A relative axis wraps at the grid edge, as Excel's does.
  ExpectRange(Parse(wb.handle, "R[-1]C", 1, 1, 0, 0), kLastRow, 0, kLastRow, 0);
  EXPECT_FALSE(Parse(wb.handle, "R[-1]C", 1).valid);

  WorkbookGuard de;
  ASSERT_EQ(fm_workbook_create_default(&de.handle), 0);
  ASSERT_EQ(fm_workbook_set_excel_profile_id(de.handle, "mac-365-de_DE"), 0);
  ExpectRange(Parse(de.handle, "Z2S3", 1), 1, 2, 1, 2);
}

TEST(FormulonCApiReference, SheetNameOutlivesTheNextParserCall) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_parsed_reference_t out{};
  ASSERT_EQ(fm_workbook_parse_reference(wb.handle, "'My Sheet'!A1", 0, 0, 0, 0, &out), 0);
  ASSERT_EQ(out.has_sheet, 1);
  // The parser's own sheet view lives in thread-local scratch the next parse reuses.
  const auto other = formulon::eval::refs_internal::parse_a1_ref("'Other one'!B2");
  ASSERT_TRUE(other.valid);
  EXPECT_STREQ(out.sheet_name, "My Sheet");
}

TEST(FormulonCApiReference, ParseRejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_parsed_reference_t out{};
  EXPECT_EQ(fm_workbook_parse_reference(nullptr, "A1", 0, 0, 0, 0, &out), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_parse_reference(wb.handle, nullptr, 0, 0, 0, 0, &out), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_parse_reference(wb.handle, "A1", 0, 0, 0, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_parse_reference(wb.handle, "A\xff", 0, 0, 0, 0, &out), kBindingUtf8);
  EXPECT_EQ(fm_workbook_parse_reference(wb.handle, "A1", 2, 0, 0, 0, &out), kInvalidArgument);
  EXPECT_EQ(fm_workbook_parse_reference(wb.handle, "A1", -1, 0, 0, 0, &out), kInvalidArgument);
  EXPECT_EQ(fm_workbook_parse_reference(wb.handle, "A1", 0, 2, 0, 0, &out), kInvalidArgument);
  // The base is checked in A1 mode too.
  EXPECT_EQ(fm_workbook_parse_reference(wb.handle, "A1", 0, 1, kLastRow + 1U, 0, &out), kInvalidArgument);
  EXPECT_EQ(fm_workbook_parse_reference(wb.handle, "RC", 1, 1, 0, kLastCol + 1U, &out), kInvalidArgument);
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
  EXPECT_EQ(fm_workbook_format_reference(nullptr, nullptr, 0, 0, 0, 0, 0, &text), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_format_reference(wb.handle, nullptr, 0, 0, 0, 0, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_format_reference(wb.handle, "S\xff", 0, 0, 0, 0, 0, &text), kBindingUtf8);
  EXPECT_EQ(fm_workbook_format_reference(wb.handle, nullptr, 0, 0, kLastRow + 1U, 0, 0, &text), kInvalidArgument);
  EXPECT_EQ(fm_workbook_format_reference(wb.handle, nullptr, 0, 0, 0, kLastCol + 1U, 0, &text), kInvalidArgument);
  EXPECT_EQ(fm_workbook_format_reference(wb.handle, nullptr, 2, 0, 1, 0, 0, &text), kInvalidArgument);
  EXPECT_EQ(fm_workbook_format_reference(wb.handle, nullptr, 0, 2, 0, 1, 0, &text), kInvalidArgument);
  EXPECT_EQ(fm_workbook_format_reference(wb.handle, nullptr, 0, 0, 0, 0, 2, &text), kInvalidArgument);
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
