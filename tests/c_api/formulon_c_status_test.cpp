//
// Status codes of the C ABI: the generated `FM_STATUS_*` defines, the names
// `fm_status_string` returns, and their agreement with the engine's own
// error enum.

#include <cctype>
#include <cstdint>
#include <cstring>
#include <string>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

// `kParserInvalidR1C1` -> `PARSER_INVALID_R1C1`: a word starts at an upper-case
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
  EXPECT_EQ(FM_STATUS_PARSER_INVALID_R1C1, static_cast<int>(formulon::FormulonErrorCode::kParserInvalidR1C1));
  EXPECT_EQ(FM_STATUS_SHEET_INDEX_OUT_OF_RANGE, static_cast<int>(formulon::FormulonErrorCode::kSheetIndexOutOfRange));
  EXPECT_EQ(FM_STATUS_BINDING_NULL_POINTER, static_cast<int>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(FM_STATUS_THREADED_COMMENT_INVALID, static_cast<int>(formulon::FormulonErrorCode::kThreadedCommentInvalid));
}

TEST(FormulonCStatus, StatusStringIsOkForZero) {
  EXPECT_STREQ(fm_status_string(0), "OK");
  EXPECT_STREQ(fm_status_string(FM_STATUS_SHEET_INDEX_OUT_OF_RANGE), "SHEET_INDEX_OUT_OF_RANGE");
  EXPECT_STREQ(fm_status_string(FM_STATUS_PARSER_INVALID_R1C1), "PARSER_INVALID_R1C1");
}

TEST(FormulonCStatus, UnknownValuesAreUnrecognizedNotUnknownError) {
  EXPECT_STREQ(fm_status_string(FM_STATUS_UNKNOWN_ERROR), "UNKNOWN_ERROR");
  for (const fm_status_t value : {-1, 10, 999, 1015, 5051 + 100, 10000, 123456, INT32_MIN, INT32_MAX}) {
    EXPECT_STREQ(fm_status_string(value), "UNRECOGNIZED_STATUS") << value;
  }
}

TEST(FormulonCStatus, EveryEngineCodeIsNamedAndNothingElseIs) {
  int named = 0;
  for (fm_status_t value = 0; value < 10000; ++value) {
    const char* engine_name = formulon::to_cstring(static_cast<formulon::FormulonErrorCode>(value));
    const std::string c_name = fm_status_string(value);
    if (c_name == "UNRECOGNIZED_STATUS") {
      EXPECT_STREQ(engine_name, "kUnknownError") << value;
    } else {
      ++named;
      EXPECT_EQ(c_name, UpperSnake(engine_name)) << value;
    }
  }
  EXPECT_EQ(named, 114);
}
