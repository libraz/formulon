//
// Stable C ABI function-catalog metadata tests.

#include <cctype>
#include <cstdint>
#include <cstring>
#include <limits>
#include <string>
#include <unordered_set>

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

TEST(FormulonCApiFunctionMetadata, RawLocaleValuesAreRejectedWithoutMutation) {
  const std::int32_t invalid_locales[] = {99, std::numeric_limits<std::int32_t>::min(),
                                          std::numeric_limits<std::int32_t>::max()};
  const fm_status_t expected = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  for (const std::int32_t raw : invalid_locales) {
    fm_function_metadata_t md{};
    md.canonical_name = "sentinel";
    EXPECT_EQ(fm_function_metadata("SUM", raw, &md), expected);
    EXPECT_EQ(md.canonical_name, nullptr);

    const char* localized = "sentinel";
    EXPECT_EQ(fm_function_localize("SUM", raw, &localized), expected);
    EXPECT_EQ(localized, nullptr);

    const char* canonical = "sentinel";
    EXPECT_EQ(fm_function_canonicalize("SUM", raw, &canonical), expected);
    EXPECT_EQ(canonical, nullptr);
  }
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

// Lazy-dispatch forms (XLOOKUP, SUMIFS, ...) and parser special forms (LET,
// LAMBDA) are recognised by the evaluator but are NOT in the eager registry.
// The catalog must still enumerate them and resolve their metadata.
TEST(FormulonCApiFunctionMetadata, EnumerationIncludesLazyAndSpecialForms) {
  std::unordered_set<std::string> names;
  const std::size_t count = fm_function_count();
  for (std::size_t i = 0; i < count; ++i) {
    const char* name = nullptr;
    ASSERT_EQ(fm_function_name_at(i, &name), 0);
    ASSERT_NE(name, nullptr);
    names.insert(name);
  }
  for (const char* expected : {"XLOOKUP", "SUMIFS", "IFERROR", "INDEX", "OFFSET", "INDIRECT", "SORT", "UNIQUE",
                               "FILTER", "LET", "LAMBDA", "VLOOKUP"}) {
    EXPECT_TRUE(names.count(expected) != 0) << expected << " missing from catalog enumeration";
  }
}

TEST(FormulonCApiFunctionMetadata, LazyAndSpecialFormsResolveMetadata) {
  for (const char* fn : {"XLOOKUP", "SUMIFS", "IFERROR", "INDEX", "OFFSET", "INDIRECT", "SORT", "UNIQUE", "FILTER",
                         "LET", "LAMBDA", "VLOOKUP"}) {
    fm_function_metadata_t md{};
    ASSERT_EQ(fm_function_metadata(fn, FM_LOCALE_EN_US, &md), 0) << fn << " did not resolve";
    ASSERT_NE(md.canonical_name, nullptr);
    EXPECT_STREQ(md.canonical_name, fn);
    // No FunctionDef -> arity is unknown: min 0, max unbounded sentinel.
    EXPECT_EQ(md.min_arity, 0U);
    EXPECT_EQ(md.max_arity, 0xFFFFFFFFU);
    EXPECT_EQ(md.availability, FM_FUNCTION_IMPLEMENTED);
  }
}

TEST(FormulonCApiFunctionMetadata, LazyFormLookupIsCaseInsensitive) {
  fm_function_metadata_t md{};
  ASSERT_EQ(fm_function_metadata("xlookup", FM_LOCALE_EN_US, &md), 0);
  EXPECT_STREQ(md.canonical_name, "XLOOKUP");
}

// Every enumerated name — eager, lazy, and special form alike — must
// resolve through the same catalog APIs. Regression guard for the
// membership split where localize / canonicalize consulted only the eager
// registry and rejected XLOOKUP / LET / SUMIFS despite enumerating them.
TEST(FormulonCApiFunctionMetadata, EveryEnumeratedNameRoundTripsAcrossAllCatalogApis) {
  const std::size_t count = fm_function_count();
  ASSERT_GT(count, 0U);
  for (std::size_t i = 0; i < count; ++i) {
    const char* name = nullptr;
    ASSERT_EQ(fm_function_name_at(i, &name), 0);
    ASSERT_NE(name, nullptr);

    // metadata
    fm_function_metadata_t md{};
    EXPECT_EQ(fm_function_metadata(name, FM_LOCALE_EN_US, &md), 0) << name << " has no metadata";

    // canonicalize: an enumerated name is already canonical, so it maps to
    // itself.
    const char* canonical = nullptr;
    ASSERT_EQ(fm_function_canonicalize(name, FM_LOCALE_EN_US, &canonical), 0) << name << " did not canonicalize";
    ASSERT_NE(canonical, nullptr);
    EXPECT_STREQ(canonical, name);

    // localize: with no alias table it falls through to the canonical name.
    const char* localized = nullptr;
    ASSERT_EQ(fm_function_localize(name, FM_LOCALE_EN_US, &localized), 0) << name << " did not localize";
    ASSERT_NE(localized, nullptr);
    EXPECT_STREQ(localized, name);
  }
}

TEST(FormulonCApiFunctionMetadata, LazyAndSpecialFormsCanonicalizeCaseInsensitively) {
  for (const char* fn : {"xlookup", "sumifs", "let", "lambda", "filter"}) {
    const char* canonical = nullptr;
    ASSERT_EQ(fm_function_canonicalize(fn, FM_LOCALE_EN_US, &canonical), 0) << fn << " did not canonicalize";
    ASSERT_NE(canonical, nullptr);
    std::string upper(fn);
    for (char& c : upper) {
      c = static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
    }
    EXPECT_EQ(std::string(canonical), upper);
  }
}
