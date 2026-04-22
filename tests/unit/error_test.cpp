// Copyright 2026 libraz. Licensed under the MIT License.
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
      {FormulonErrorCode::kParserUnexpectedToken, "kParserUnexpectedToken"},
      {FormulonErrorCode::kEvalCircularReference, "kEvalCircularReference"},
      {FormulonErrorCode::kFnNotRegistered, "kFnNotRegistered"},
      {FormulonErrorCode::kGraphCycleDetected, "kGraphCycleDetected"},
      {FormulonErrorCode::kIoZipSlip, "kIoZipSlip"},
      {FormulonErrorCode::kCryptoBadPassword, "kCryptoBadPassword"},
      {FormulonErrorCode::kBindingInvalidHandle, "kBindingInvalidHandle"},
      {FormulonErrorCode::kPivotSourceInvalid, "kPivotSourceInvalid"},
      {FormulonErrorCode::kPrintLayoutConvergence, "kPrintLayoutConvergence"},
  };
  for (const auto& s : samples) {
    EXPECT_STREQ(s.name, to_cstring(s.code)) << "code=" << static_cast<int>(s.code);
  }
}

TEST(ErrorTest, ErrorCodesAreWithinDocumentedRanges) {
  // Band boundaries per backup/plans/23-error-codes.md §23.3.
  static_assert(static_cast<int32_t>(FormulonErrorCode::kOk) == 0, "kOk must be zero");
  static_assert(static_cast<int32_t>(FormulonErrorCode::kUnknownError) == 1, "kUnknownError must be 1");
  static_assert(static_cast<int32_t>(FormulonErrorCode::kParserUnexpectedToken) == 1000, "parser band starts at 1000");
  static_assert(static_cast<int32_t>(FormulonErrorCode::kEvalStackOverflow) == 2000, "evaluator band starts at 2000");
  static_assert(static_cast<int32_t>(FormulonErrorCode::kFnNotRegistered) == 3000, "functions band starts at 3000");
  static_assert(static_cast<int32_t>(FormulonErrorCode::kGraphCycleDetected) == 4000, "graph band starts at 4000");
  static_assert(static_cast<int32_t>(FormulonErrorCode::kIoFileNotFound) == 5000, "I/O band starts at 5000");
  static_assert(static_cast<int32_t>(FormulonErrorCode::kCryptoAgileNotSupported) == 6000,
                "crypto band starts at 6000");
  static_assert(static_cast<int32_t>(FormulonErrorCode::kBindingInvalidHandle) == 7000, "bindings band starts at 7000");
  static_assert(static_cast<int32_t>(FormulonErrorCode::kPivotSourceInvalid) == 8000, "pivot band starts at 8000");
  static_assert(static_cast<int32_t>(FormulonErrorCode::kPrintFontMetricsMissing) == 9000, "print band starts at 9000");
  SUCCEED();
}

TEST(ErrorTest, UnknownEnumValueFallsBackToUnknown) {
  const auto bogus = static_cast<FormulonErrorCode>(123456);
  EXPECT_STREQ("kUnknownError", to_cstring(bogus));
}

}  // namespace
}  // namespace formulon
