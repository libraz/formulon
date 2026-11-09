//
// Unit tests for Error / FormulonErrorCode utilities.

#include "utils/error.h"

#include <cstdint>
#include <string>

#include "gtest/gtest.h"

namespace formulon {
namespace {

TEST(ErrorTest, MakeErrorStoresFields) {
  Error err = make_error(FormulonErrorCode::kInvalidArgument, "bad argument");
  EXPECT_EQ(FormulonErrorCode::kInvalidArgument, err.code);
  EXPECT_EQ("bad argument", err.message);
  EXPECT_TRUE(err.context.empty());
}

TEST(ErrorTest, MakeErrorWithContextPreservesContext) {
  Error err = make_error(FormulonErrorCode::kIoFileNotFound, "open failed", "path=workbook.xlsx");
  EXPECT_EQ(FormulonErrorCode::kIoFileNotFound, err.code);
  EXPECT_EQ("open failed", err.message);
  EXPECT_EQ("path=workbook.xlsx", err.context);
}

TEST(ErrorTest, ErrorCodeNameRoundTrip) {
  struct Sample {
    FormulonErrorCode code;
    const char* name;
  };
  const Sample samples[] = {
      {FormulonErrorCode::kOk, "kOk"},
      {FormulonErrorCode::kAlreadyExists, "kAlreadyExists"},
      {FormulonErrorCode::kParserUnexpectedToken, "kParserUnexpectedToken"},
      {FormulonErrorCode::kEvalPivotInvalid, "kEvalPivotInvalid"},
      {FormulonErrorCode::kFnOverflow, "kFnOverflow"},
      {FormulonErrorCode::kGraphRecalcReentrant, "kGraphRecalcReentrant"},
      {FormulonErrorCode::kIoZipSlip, "kIoZipSlip"},
      {FormulonErrorCode::kSecResourceLimit, "kSecResourceLimit"},
      {FormulonErrorCode::kBindingInvalidHandle, "kBindingInvalidHandle"},
      {FormulonErrorCode::kAutoFilterInvalid, "kAutoFilterInvalid"},
      {FormulonErrorCode::kPrintPageCountOverflow, "kPrintPageCountOverflow"},
  };
  for (const auto& s : samples) {
    EXPECT_STREQ(s.name, to_cstring(s.code)) << "code=" << static_cast<int>(s.code);
  }
}

TEST(ErrorTest, ErrorCodesAreWithinDocumentedRanges) {
  // Each code lies in its module's 1000-wide band.
  constexpr auto band = [](FormulonErrorCode code) { return static_cast<int32_t>(code) / 1000; };
  static_assert(static_cast<int32_t>(FormulonErrorCode::kOk) == 0, "kOk must be zero");
  static_assert(band(FormulonErrorCode::kPreconditionFailed) == 0, "general band is 0-999");
  static_assert(band(FormulonErrorCode::kParserUnexpectedToken) == 1, "parser band is 1000-1999");
  static_assert(band(FormulonErrorCode::kEvalPivotMissing) == 2, "evaluator band is 2000-2999");
  static_assert(band(FormulonErrorCode::kFnOverflow) == 3, "functions band is 3000-3999");
  static_assert(band(FormulonErrorCode::kGraphScheduleFailed) == 4, "graph band is 4000-4999");
  static_assert(band(FormulonErrorCode::kIoFileNotFound) == 5, "I/O band is 5000-5999");
  static_assert(band(FormulonErrorCode::kSecResourceLimit) == 6, "security band is 6000-6999");
  static_assert(band(FormulonErrorCode::kBindingInvalidHandle) == 7, "bindings band is 7000-7999");
  static_assert(band(FormulonErrorCode::kDataTableInvalid) == 8, "pivot band is 8000-8999");
  static_assert(band(FormulonErrorCode::kPrintInvalidArea) == 9, "print band is 9000-9999");
  SUCCEED();
}

TEST(ErrorTest, UnknownEnumValueFallsBackToUnknown) {
  const auto bogus = static_cast<FormulonErrorCode>(123456);
  EXPECT_STREQ("kUnknownError", to_cstring(bogus));
}

}  // namespace
}  // namespace formulon
