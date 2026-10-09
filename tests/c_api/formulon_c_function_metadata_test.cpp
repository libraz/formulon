//
// Stable C ABI function-catalog metadata tests.

#include <cctype>
#include <cstdint>
#include <cstring>
#include <string>
#include <unordered_set>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

TEST(FormulonCApiFunctionMetadata, KnownFunctionResolves) {
  fm_function_metadata_t md{};
  ASSERT_EQ(fm_function_metadata("SUM", &md), 0);
  ASSERT_NE(md.canonical_name, nullptr);
  EXPECT_STREQ(md.canonical_name, "SUM");
  EXPECT_EQ(md.min_arity, 1U);
  // SUM is variadic.
  EXPECT_EQ(md.max_arity, 0xFFFFFFFFU);
  EXPECT_EQ(md.availability, FM_FUNCTION_IMPLEMENTED);
  // Display text is host-supplied; the engine always reports NULL.
  EXPECT_EQ(md.signature_template, nullptr);
  EXPECT_EQ(md.description, nullptr);
}

TEST(FormulonCApiFunctionMetadata, LookupIsCaseInsensitive) {
  fm_function_metadata_t md{};
  ASSERT_EQ(fm_function_metadata("sum", &md), 0);
  EXPECT_STREQ(md.canonical_name, "SUM");
  ASSERT_EQ(fm_function_metadata("SuM", &md), 0);
  EXPECT_STREQ(md.canonical_name, "SUM");
}

TEST(FormulonCApiFunctionMetadata, UnknownFunctionReturnsInvalidArgument) {
  fm_function_metadata_t md{};
  fm_status_t rc = fm_function_metadata("NOT_A_FUNCTION", &md);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiFunctionMetadata, AvailabilityDistinguishesUnavailableStubs) {
  fm_function_metadata_t md{};
  ASSERT_EQ(fm_function_metadata("WEBSERVICE", &md), 0);
  EXPECT_STREQ(md.canonical_name, "WEBSERVICE");
  EXPECT_EQ(md.availability, FM_FUNCTION_UNAVAILABLE_STUB);

  ASSERT_EQ(fm_function_metadata("CUBEVALUE", &md), 0);
  EXPECT_EQ(md.availability, FM_FUNCTION_UNAVAILABLE_STUB);
}

TEST(FormulonCApiFunctionMetadata, AvailabilityDistinguishesNonStubSpecialCases) {
  fm_function_metadata_t md{};
  ASSERT_EQ(fm_function_metadata("FILTERXML", &md), 0);
  EXPECT_EQ(md.availability, FM_FUNCTION_IMPLEMENTED);

  ASSERT_EQ(fm_function_metadata("INFO", &md), 0);
  EXPECT_EQ(md.availability, FM_FUNCTION_ENVIRONMENT_BOUND);

  ASSERT_EQ(fm_function_metadata("SUM", &md), 0);
  EXPECT_EQ(md.availability, FM_FUNCTION_IMPLEMENTED);
}

TEST(FormulonCApiFunctionMetadata, NullArgsReturnBindingNullPointer) {
  fm_function_metadata_t md{};
  EXPECT_EQ(fm_function_metadata(nullptr, &md),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_function_metadata("SUM", nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiFunctionMetadata, UnknownProfileIsRejectedWithoutMutation) {
  const fm_status_t invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  for (const char* id : {"", "de_DE", "mac-365-xx_XX", "win-365-en-US"}) {
    const char* localized = "sentinel";
    EXPECT_EQ(fm_function_localize("SUM", id, &localized), invalid) << id;
    EXPECT_EQ(localized, nullptr);
    EXPECT_NE(std::strstr(fm_last_error_context(), "profile_id="), nullptr) << id;

    const char* canonical = "sentinel";
    EXPECT_EQ(fm_function_canonicalize("SUM", id, &canonical), invalid) << id;
    EXPECT_EQ(canonical, nullptr);
  }
}

TEST(FormulonCApiFunctionMetadata, NullProfileReturnsBindingNullPointer) {
  const fm_status_t null_ptr = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
  const char* out = nullptr;
  EXPECT_EQ(fm_function_localize("SUM", nullptr, &out), null_ptr);
  EXPECT_EQ(fm_function_canonicalize("SUM", nullptr, &out), null_ptr);
  EXPECT_EQ(fm_function_localize(nullptr, "mac-365-de_DE", &out), null_ptr);
  EXPECT_EQ(fm_function_canonicalize("SUM", "mac-365-de_DE", nullptr), null_ptr);
}

TEST(FormulonCApiFunctionMetadata, GermanAndFrenchShowLocalizedNames) {
  const char* out = nullptr;
  for (const char* id : {"mac-365-de_DE", "win-365-de_DE"}) {
    ASSERT_EQ(fm_function_localize("sum", id, &out), 0);
    EXPECT_STREQ(out, "SUMME");
    ASSERT_EQ(fm_function_canonicalize("SUMME", id, &out), 0);
    EXPECT_STREQ(out, "SUM");
    // The canonical spelling stays accepted as a fallback.
    ASSERT_EQ(fm_function_canonicalize("sum", id, &out), 0);
    EXPECT_STREQ(out, "SUM");
  }
  ASSERT_EQ(fm_function_localize("XLOOKUP", "mac-365-fr_FR", &out), 0);
  EXPECT_STREQ(out, "RECHERCHEX");
  ASSERT_EQ(fm_function_canonicalize("SOMME", "mac-365-fr_FR", &out), 0);
  EXPECT_STREQ(out, "SUM");
  ASSERT_EQ(fm_function_canonicalize("TRIM", "mac-365-fr_FR", &out), 0);
  EXPECT_STREQ(out, "MIRR");
  // A localized spelling of another locale is not a function here.
  EXPECT_NE(fm_function_canonicalize("SUMME", "mac-365-fr_FR", &out), 0);
}

TEST(FormulonCApiFunctionMetadata, OtherLocalesShowCanonicalNames) {
  const char* out = nullptr;
  for (const char* id : {"mac-365-ja_JP", "win-365-ja_JP", "mac-365-en_US", "win-365-en_US", "mac-365-zh_CN",
                         "mac-365-ko_KR", "mac-365-th_TH"}) {
    ASSERT_EQ(fm_function_localize("SUM", id, &out), 0) << id;
    EXPECT_STREQ(out, "SUM") << id;
    ASSERT_EQ(fm_function_canonicalize("sum", id, &out), 0) << id;
    EXPECT_STREQ(out, "SUM") << id;
  }
  for (const char* id : {"mac-365-en_US", "mac-365-zh_CN", "mac-365-ko_KR", "mac-365-th_TH"}) {
    ASSERT_EQ(fm_function_localize("DOLLAR", id, &out), 0) << id;
    EXPECT_STREQ(out, "DOLLAR") << id;
  }
}

// ja-JP renames only the three functions its name column lists.
TEST(FormulonCApiFunctionMetadata, JapaneseRenamesTheListedFunctions) {
  const char* out = nullptr;
  ASSERT_EQ(fm_function_localize("DOLLAR", "mac-365-ja_JP", &out), 0);
  EXPECT_STREQ(out, "YEN");
  ASSERT_EQ(fm_function_canonicalize("YEN", "win-365-ja_JP", &out), 0);
  EXPECT_STREQ(out, "DOLLAR");
  ASSERT_EQ(fm_function_localize("SUM", "mac-365-ja_JP", &out), 0);
  EXPECT_STREQ(out, "SUM");
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
    ASSERT_EQ(fm_function_metadata(fn, &md), 0) << fn << " did not resolve";
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
  ASSERT_EQ(fm_function_metadata("xlookup", &md), 0);
  EXPECT_STREQ(md.canonical_name, "XLOOKUP");
}

// Every enumerated name — eager, lazy, and special form alike — must
// resolve through the same catalog APIs. Regression guard for the
// membership split where localize / canonicalize consulted only the eager
// registry and rejected XLOOKUP / LET / SUMIFS despite enumerating them.
//
// Swept across locales with no name column, because metadata is
// locale-invariant and a function-text table growing inside the engine would
// break the merge contract in `docs/function-metadata-schema.md`.
TEST(FormulonCApiFunctionMetadata, EveryEnumeratedNameRoundTripsAcrossAllCatalogApis) {
  const std::size_t count = fm_function_count();
  ASSERT_GT(count, 0U);
  for (const char* profile : {"win-365-en_US", "mac-365-zh_CN"}) {
    for (std::size_t i = 0; i < count; ++i) {
      const char* name = nullptr;
      ASSERT_EQ(fm_function_name_at(i, &name), 0);
      ASSERT_NE(name, nullptr);

      // metadata
      fm_function_metadata_t md{};
      EXPECT_EQ(fm_function_metadata(name, &md), 0) << name << " has no metadata";
      // Display text is host-supplied, never engine-owned. Asserted for
      // the whole catalog rather than one sample: the eager and the
      // lazy / special-form branches populate the result separately, so
      // a single function only covers one of them.
      EXPECT_EQ(md.signature_template, nullptr) << name << " carries engine-side signature text";
      EXPECT_EQ(md.description, nullptr) << name << " carries engine-side description text";

      // canonicalize: an enumerated name is already canonical, so it maps to
      // itself.
      const char* canonical = nullptr;
      ASSERT_EQ(fm_function_canonicalize(name, profile, &canonical), 0) << name << " did not canonicalize";
      ASSERT_NE(canonical, nullptr);
      EXPECT_STREQ(canonical, name);

      // localize: with no alias table it falls through to the canonical name.
      const char* localized = nullptr;
      ASSERT_EQ(fm_function_localize(name, profile, &localized), 0) << name << " did not localize";
      ASSERT_NE(localized, nullptr);
      EXPECT_STREQ(localized, name);
    }
  }
}

TEST(FormulonCApiFunctionMetadata, LazyAndSpecialFormsCanonicalizeCaseInsensitively) {
  for (const char* fn : {"xlookup", "sumifs", "let", "lambda", "filter"}) {
    const char* canonical = nullptr;
    ASSERT_EQ(fm_function_canonicalize(fn, "win-365-en_US", &canonical), 0) << fn << " did not canonicalize";
    ASSERT_NE(canonical, nullptr);
    std::string upper(fn);
    for (char& c : upper) {
      c = static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
    }
    EXPECT_EQ(std::string(canonical), upper);
  }
}
