// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for RETURN_IF_ERROR / ASSIGN_OR_RETURN.

#include "utils/status_macros.h"

#include <string>
#include <utility>

#include "gtest/gtest.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace {

Expected<int> FetchInt(bool ok) {
  if (ok)
    return 7;
  return make_error(FormulonErrorCode::kInvalidArgument, "fetch failed");
}

Expected<std::string> BuildLabel(bool ok_a, bool ok_b) {
  ASSIGN_OR_RETURN(int a, FetchInt(ok_a));
  ASSIGN_OR_RETURN(int b, FetchInt(ok_b));
  return std::string("a=") + std::to_string(a) + ",b=" + std::to_string(b);
}

Expected<void> Validate(bool ok) {
  if (ok)
    return Expected<void>::Ok();
  return make_error(FormulonErrorCode::kPreconditionFailed, "no");
}

Expected<int> RunAfterValidation(bool ok) {
  RETURN_IF_ERROR(Validate(ok));
  return 1;
}

TEST(StatusMacrosTest, ReturnIfErrorPropagatesError) {
  Expected<int> r = RunAfterValidation(false);
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(FormulonErrorCode::kPreconditionFailed, r.error().code);
}

TEST(StatusMacrosTest, ReturnIfErrorFallsThroughOnOk) {
  Expected<int> r = RunAfterValidation(true);
  ASSERT_TRUE(r.has_value());
  EXPECT_EQ(1, r.value());
}

TEST(StatusMacrosTest, AssignOrReturnBindsValue) {
  Expected<std::string> r = BuildLabel(true, true);
  ASSERT_TRUE(r.has_value());
  EXPECT_EQ("a=7,b=7", r.value());
}

TEST(StatusMacrosTest, AssignOrReturnPropagatesError) {
  Expected<std::string> r = BuildLabel(true, false);
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(FormulonErrorCode::kInvalidArgument, r.error().code);
  EXPECT_EQ("fetch failed", r.error().message);
}

// Exercises per-call-site uniqueness of the internal tmp variable. If
// `__COUNTER__` / `__LINE__` concatenation broke, this translation unit
// would fail to compile with "redeclaration" errors.
Expected<int> SumThree() {
  ASSIGN_OR_RETURN(int a, FetchInt(true));
  ASSIGN_OR_RETURN(int b, FetchInt(true));
  ASSIGN_OR_RETURN(int c, FetchInt(true));
  return a + b + c;
}

TEST(StatusMacrosTest, AssignOrReturnWorksInSameScopeWithMultipleUses) {
  Expected<int> r = SumThree();
  ASSERT_TRUE(r.has_value());
  EXPECT_EQ(21, r.value());
}

}  // namespace
}  // namespace formulon
