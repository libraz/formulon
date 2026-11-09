//
// Status codes of the C ABI: the generated `FM_STATUS_*` defines, the names
// `fm_get_status_name` returns, and their agreement with the engine's own
// error enum.

#include <cctype>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

// `kParserInvalidReference` -> `PARSER_INVALID_REFERENCE`: a word starts at an upper-case
// letter that follows a lower-case one; digits never start a word.
std::string UpperSnake(const char* enumerator_name) {
  std::string out;
  const char* p = enumerator_name + 1;  // skip the `k` prefix
  for (bool prev_lower = false; *p != '\0'; ++p) {
    const bool upper = std::isupper(static_cast<unsigned char>(*p)) != 0;
    if (upper && prev_lower) {
      out += '_';
    }
    out += static_cast<char>(std::toupper(static_cast<unsigned char>(*p)));
    prev_lower = std::islower(static_cast<unsigned char>(*p)) != 0;
  }
  return out;
}

}  // namespace

TEST(FormulonCStatus, DefinesCarryTheEngineValues) {
  EXPECT_EQ(FM_STATUS_OK, static_cast<int>(formulon::FormulonErrorCode::kOk));
  EXPECT_EQ(FM_STATUS_INVALID_ARGUMENT, static_cast<int>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(FM_STATUS_PARSER_INVALID_REFERENCE, static_cast<int>(formulon::FormulonErrorCode::kParserInvalidReference));
  EXPECT_EQ(FM_STATUS_ALREADY_EXISTS, static_cast<int>(formulon::FormulonErrorCode::kAlreadyExists));
  EXPECT_EQ(FM_STATUS_PRECONDITION_FAILED, static_cast<int>(formulon::FormulonErrorCode::kPreconditionFailed));
  EXPECT_EQ(FM_STATUS_SHEET_INDEX_OUT_OF_RANGE, static_cast<int>(formulon::FormulonErrorCode::kSheetIndexOutOfRange));
  EXPECT_EQ(FM_STATUS_BINDING_NULL_POINTER, static_cast<int>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(FM_STATUS_THREADED_COMMENT_INVALID, static_cast<int>(formulon::FormulonErrorCode::kThreadedCommentInvalid));
}

TEST(FormulonCStatus, StatusStringIsOkForZero) {
  EXPECT_STREQ(fm_get_status_name(0), "OK");
  EXPECT_STREQ(fm_get_status_name(FM_STATUS_SHEET_INDEX_OUT_OF_RANGE), "SHEET_INDEX_OUT_OF_RANGE");
  EXPECT_STREQ(fm_get_status_name(FM_STATUS_PARSER_INVALID_REFERENCE), "PARSER_INVALID_REFERENCE");
  EXPECT_STREQ(fm_get_status_name(FM_STATUS_ALREADY_EXISTS), "ALREADY_EXISTS");
}

// Codes no entry point returns are not part of the header.
#ifdef FM_STATUS_UNKNOWN_ERROR
#error "FM_STATUS_UNKNOWN_ERROR is no longer a status code"
#endif
#ifdef FM_STATUS_CANNOT_REMOVE_LAST_SHEET
#error "FM_STATUS_CANNOT_REMOVE_LAST_SHEET is no longer a status code"
#endif
#ifdef FM_STATUS_GRAPH_ITERATION_ABORTED
#error "FM_STATUS_GRAPH_ITERATION_ABORTED is no longer a status code"
#endif

TEST(FormulonCStatus, UnknownValuesAreUnrecognized) {
  for (const fm_status_t value : {-1, 1, 10, 999, 1015, 5052, 5051 + 100, 10000, 123456, INT32_MIN, INT32_MAX}) {
    EXPECT_STREQ(fm_get_status_name(value), "UNRECOGNIZED_STATUS") << value;
  }
}

TEST(FormulonCStatus, EveryEngineCodeIsNamedAndNothingElseIs) {
  int named = 0;
  for (fm_status_t value = 0; value < 10000; ++value) {
    const char* engine_name = formulon::to_cstring(static_cast<formulon::FormulonErrorCode>(value));
    const std::string c_name = fm_get_status_name(value);
    if (c_name == "UNRECOGNIZED_STATUS") {
      EXPECT_STREQ(engine_name, "kUnknownError") << value;
    } else {
      ++named;
      EXPECT_EQ(c_name, UpperSnake(engine_name)) << value;
    }
  }
  EXPECT_EQ(named, 60);
}

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

}  // namespace

// A name or id another element already holds is ALREADY_EXISTS on every
// surface; a malformed name stays an argument refusal.
TEST(FormulonCStatus, TakenNamesAndIdsAreAlreadyExists) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t index = 0;
  EXPECT_EQ(fm_workbook_add_sheet(wb.handle, "sheet1", &index), FM_STATUS_ALREADY_EXISTS);
  EXPECT_EQ(fm_workbook_add_sheet(wb.handle, "bad:name", &index), FM_STATUS_INVALID_SHEET_NAME);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Other", &index), 0);
  EXPECT_EQ(fm_workbook_rename_sheet(wb.handle, index, "SHEET1"), FM_STATUS_ALREADY_EXISTS);
  EXPECT_EQ(fm_workbook_copy_sheet(wb.handle, 0, 0, "Other", &index), FM_STATUS_ALREADY_EXISTS);

  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Region"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 1, "Amount"), 0);
  fm_table_spec_t table = FM_STRUCT_INIT(fm_table_spec_t);
  table.range = fm_cell_rect_t{0, 0, 2, 1};
  table.name = "Sales";
  std::size_t table_index = 0;
  ASSERT_EQ(fm_sheet_add_table(wb.handle, 0, &table, &table_index), 0) << fm_get_last_error_message();
  const fm_defined_name_t taken{sizeof(fm_defined_name_t), "Sales", "1", -1, nullptr, 0};
  EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, &taken), FM_STATUS_ALREADY_EXISTS);
  const fm_defined_name_t malformed{sizeof(fm_defined_name_t), "1A", "1", -1, nullptr, 0};
  EXPECT_EQ(fm_workbook_set_defined_name(wb.handle, &malformed), FM_STATUS_INVALID_ARGUMENT);

  fm_person_t person = FM_STRUCT_INIT(fm_person_t);
  person.id = "{0A1B2C3D-0000-4000-8000-000000000001}";
  person.display_name = "Alice";
  ASSERT_EQ(fm_workbook_add_person(wb.handle, &person), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_workbook_add_person(wb.handle, &person), FM_STATUS_ALREADY_EXISTS);

  const fm_cell_rect_t area{4, 4, 5, 5};
  fm_protected_range_t range = FM_STRUCT_INIT(fm_protected_range_t);
  range.name = "Edit";
  range.ranges = &area;
  range.range_count = 1;
  std::size_t range_index = 0;
  ASSERT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &range, &range_index), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_add_protected_range(wb.handle, 0, &range, &range_index), FM_STATUS_ALREADY_EXISTS);
}
