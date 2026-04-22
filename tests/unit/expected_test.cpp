// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for formulon::Expected<T, E>.

#include "utils/expected.h"

#include <memory>
#include <string>
#include <utility>

#include "gtest/gtest.h"
#include "utils/error.h"

namespace formulon {
namespace {

Expected<int> MakeOkInt(int v) {
  return v;
}
Expected<int> MakeErrInt(FormulonErrorCode code, std::string msg) {
  return make_error(code, std::move(msg));
}

TEST(ExpectedTest, ValueHeldOnSuccess) {
  Expected<int> e = 42;
  ASSERT_TRUE(e.has_value());
  EXPECT_EQ(42, e.value());
}

TEST(ExpectedTest, ErrorHeldOnFailure) {
  auto err = make_error(FormulonErrorCode::kInvalidArgument, "bad");
  Expected<int> e = err;
  ASSERT_FALSE(e.has_value());
  EXPECT_EQ(FormulonErrorCode::kInvalidArgument, e.error().code);
  EXPECT_EQ("bad", e.error().message);
}

TEST(ExpectedTest, BoolConversionReflectsState) {
  Expected<int> ok = 1;
  Expected<int> err = make_error(FormulonErrorCode::kUnknownError, "x");
  EXPECT_TRUE(static_cast<bool>(ok));
  EXPECT_FALSE(static_cast<bool>(err));
}

TEST(ExpectedTest, TakeTransfersOwnership) {
  Expected<std::unique_ptr<int>> e = std::make_unique<int>(7);
  ASSERT_TRUE(e.has_value());
  std::unique_ptr<int> moved = e.take();
  ASSERT_NE(nullptr, moved);
  EXPECT_EQ(7, *moved);
}

TEST(ExpectedTest, MapTransformsValue) {
  Expected<int> e = 3;
  Expected<int> doubled = e.map([](const int& v) { return v * 2; });
  ASSERT_TRUE(doubled.has_value());
  EXPECT_EQ(6, doubled.value());
}

TEST(ExpectedTest, MapPreservesError) {
  Expected<int> e = make_error(FormulonErrorCode::kInvalidArgument, "bad");
  Expected<int> mapped = e.map([](const int& v) { return v + 1; });
  ASSERT_FALSE(mapped.has_value());
  EXPECT_EQ(FormulonErrorCode::kInvalidArgument, mapped.error().code);
}

TEST(ExpectedTest, AndThenChainsSuccess) {
  Expected<int> e = 5;
  auto result =
      e.and_then([](const int& v) -> Expected<std::string> { return std::string("val=") + std::to_string(v); });
  ASSERT_TRUE(result.has_value());
  EXPECT_EQ("val=5", result.value());
}

TEST(ExpectedTest, AndThenPropagatesError) {
  Expected<int> e = make_error(FormulonErrorCode::kNotFound, "missing");
  auto result =
      e.and_then([](const int& v) -> Expected<std::string> { return std::string("val=") + std::to_string(v); });
  ASSERT_FALSE(result.has_value());
  EXPECT_EQ(FormulonErrorCode::kNotFound, result.error().code);
}

TEST(ExpectedTest, OrElseRecovers) {
  Expected<int> e = make_error(FormulonErrorCode::kNotFound, "missing");
  Expected<int> recovered = e.or_else([](const Error&) -> Expected<int> { return 99; });
  ASSERT_TRUE(recovered.has_value());
  EXPECT_EQ(99, recovered.value());
}

TEST(ExpectedTest, OrElsePassesThroughValue) {
  Expected<int> e = 11;
  Expected<int> recovered = e.or_else([](const Error&) -> Expected<int> { return 99; });
  ASSERT_TRUE(recovered.has_value());
  EXPECT_EQ(11, recovered.value());
}

TEST(ExpectedTest, VoidOk) {
  Expected<void> e = Expected<void>::Ok();
  ASSERT_TRUE(e.has_value());
  EXPECT_TRUE(static_cast<bool>(e));
}

TEST(ExpectedTest, VoidErr) {
  Expected<void> e = Expected<void>::Err(make_error(FormulonErrorCode::kInternalError, "boom"));
  ASSERT_FALSE(e.has_value());
  EXPECT_EQ(FormulonErrorCode::kInternalError, e.error().code);
  EXPECT_EQ("boom", e.error().message);
}

TEST(ExpectedTest, FactoryOkAndErr) {
  Expected<int> ok = Expected<int>::Ok(5);
  Expected<int> err = Expected<int>::Err(make_error(FormulonErrorCode::kUnknownError, "x"));
  EXPECT_TRUE(ok.has_value());
  EXPECT_FALSE(err.has_value());
  EXPECT_EQ(5, ok.value());
}

#if GTEST_HAS_DEATH_TEST
TEST(ExpectedDeathTest, ValueOnErrorAborts) {
  Expected<int> e = make_error(FormulonErrorCode::kInvalidArgument, "bad");
  EXPECT_DEATH(
      {
        // Use the result so the compiler cannot elide the call.
        volatile int v = e.value();
        (void)v;
      },
      ".*");
}

TEST(ExpectedDeathTest, ErrorOnValueAborts) {
  Expected<int> e = 1;
  EXPECT_DEATH(
      {
        const Error& err = e.error();
        (void)err;
      },
      ".*");
}
#endif  // GTEST_HAS_DEATH_TEST

// Helper used below.
Expected<int> UseMakers(bool ok) {
  if (ok)
    return MakeOkInt(1);
  return MakeErrInt(FormulonErrorCode::kInternalError, "nope");
}

TEST(ExpectedTest, HelperFunctionsRoundTrip) {
  EXPECT_TRUE(UseMakers(true).has_value());
  EXPECT_FALSE(UseMakers(false).has_value());
}

}  // namespace
}  // namespace formulon
