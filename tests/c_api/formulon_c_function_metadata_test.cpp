// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Stable C ABI function-catalog metadata tests.

#include <cstring>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

TEST(FormulonCApiFunctionMetadata, KnownFunctionResolves) {
  fm_function_metadata_t md{};
  ASSERT_EQ(fm_function_metadata("SUM", FM_LOCALE_EN_US, &md), 0);
  ASSERT_NE(md.canonical_name, nullptr);
  EXPECT_STREQ(md.canonical_name, "SUM");
  EXPECT_EQ(md.min_arity, 1U);
  // SUM is variadic.
  EXPECT_EQ(md.max_arity, 0xFFFFFFFFU);
  EXPECT_EQ(md.availability, FM_FUNCTION_IMPLEMENTED);
  // signature_template / description are not yet populated.
  EXPECT_EQ(md.signature_template, nullptr);
  EXPECT_EQ(md.description, nullptr);
}

TEST(FormulonCApiFunctionMetadata, LookupIsCaseInsensitive) {
  fm_function_metadata_t md{};
  ASSERT_EQ(fm_function_metadata("sum", FM_LOCALE_EN_US, &md), 0);
  EXPECT_STREQ(md.canonical_name, "SUM");
  ASSERT_EQ(fm_function_metadata("SuM", FM_LOCALE_EN_US, &md), 0);
  EXPECT_STREQ(md.canonical_name, "SUM");
}

TEST(FormulonCApiFunctionMetadata, UnknownFunctionReturnsInvalidArgument) {
  fm_function_metadata_t md{};
  fm_status_t rc = fm_function_metadata("NOT_A_FUNCTION", FM_LOCALE_EN_US, &md);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiFunctionMetadata, AvailabilityDistinguishesUnavailableStubs) {
  fm_function_metadata_t md{};
  ASSERT_EQ(fm_function_metadata("WEBSERVICE", FM_LOCALE_EN_US, &md), 0);
  EXPECT_STREQ(md.canonical_name, "WEBSERVICE");
  EXPECT_EQ(md.availability, FM_FUNCTION_UNAVAILABLE_STUB);

  ASSERT_EQ(fm_function_metadata("CUBEVALUE", FM_LOCALE_EN_US, &md), 0);
  EXPECT_EQ(md.availability, FM_FUNCTION_UNAVAILABLE_STUB);
}

TEST(FormulonCApiFunctionMetadata, AvailabilityDistinguishesNonStubSpecialCases) {
  fm_function_metadata_t md{};
  ASSERT_EQ(fm_function_metadata("FILTERXML", FM_LOCALE_EN_US, &md), 0);
  EXPECT_EQ(md.availability, FM_FUNCTION_IMPLEMENTED);

  ASSERT_EQ(fm_function_metadata("INFO", FM_LOCALE_EN_US, &md), 0);
  EXPECT_EQ(md.availability, FM_FUNCTION_ENVIRONMENT_BOUND);

  ASSERT_EQ(fm_function_metadata("SUM", FM_LOCALE_EN_US, &md), 0);
  EXPECT_EQ(md.availability, FM_FUNCTION_IMPLEMENTED);
}

TEST(FormulonCApiFunctionMetadata, NullArgsReturnBindingNullPointer) {
  fm_function_metadata_t md{};
  EXPECT_EQ(fm_function_metadata(nullptr, FM_LOCALE_EN_US, &md),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_function_metadata("SUM", FM_LOCALE_EN_US, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiFunctionMetadata, FunctionCountIsPositive) {
  EXPECT_GT(fm_function_count(), 100U);
}

TEST(FormulonCApiFunctionMetadata, FunctionNamesAreSortedAndComplete) {
  const std::size_t count = fm_function_count();
  ASSERT_GT(count, 0U);
  const char* prev = nullptr;
  for (std::size_t i = 0; i < count; ++i) {
    const char* name = nullptr;
    ASSERT_EQ(fm_function_name_at(i, &name), 0);
    ASSERT_NE(name, nullptr);
    if (prev != nullptr) {
      EXPECT_LT(std::strcmp(prev, name), 0) << "names not sorted at idx=" << i;
    }
    prev = name;
  }
}

TEST(FormulonCApiFunctionMetadata, NameAtOutOfRangeReturnsInvalidArgument) {
  const char* out = nullptr;
  fm_status_t rc = fm_function_name_at(static_cast<std::size_t>(-1), &out);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}
