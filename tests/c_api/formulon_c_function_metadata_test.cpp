//
// Stable C ABI function-catalog metadata tests.

#include <cctype>
#include <cstdint>
#include <cstring>
#include <string>
#include <unordered_set>

#include "c_api/formulon_c.h"
#include "formulon_c_count_of.h"
#include "gtest/gtest.h"
#include "utils/error.h"

TEST(FormulonCApiFunctionMetadata, KnownFunctionResolves) {
  int32_t present = 0;
  fm_function_metadata_t md = FM_STRUCT_INIT(fm_function_metadata_t);
  ASSERT_EQ(fm_get_function_metadata("SUM", &present, &md), 0);
  EXPECT_EQ(present, 1);
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
  int32_t present = 0;
  fm_function_metadata_t md = FM_STRUCT_INIT(fm_function_metadata_t);
  ASSERT_EQ(fm_get_function_metadata("sum", &present, &md), 0);
  EXPECT_STREQ(md.canonical_name, "SUM");
  ASSERT_EQ(fm_get_function_metadata("SuM", &present, &md), 0);
  EXPECT_STREQ(md.canonical_name, "SUM");
}

TEST(FormulonCApiFunctionMetadata, UnknownFunctionIsAbsent) {
  int32_t present = 1;
  fm_function_metadata_t md = FM_STRUCT_INIT(fm_function_metadata_t);
  ASSERT_EQ(fm_get_function_metadata("NOT_A_FUNCTION", &present, &md), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(md.canonical_name, nullptr);
}

TEST(FormulonCApiFunctionMetadata, AvailabilityDistinguishesUnavailableStubs) {
  int32_t present = 0;
  fm_function_metadata_t md = FM_STRUCT_INIT(fm_function_metadata_t);
  ASSERT_EQ(fm_get_function_metadata("WEBSERVICE", &present, &md), 0);
  EXPECT_STREQ(md.canonical_name, "WEBSERVICE");
  EXPECT_EQ(md.availability, FM_FUNCTION_UNAVAILABLE_STUB);

  ASSERT_EQ(fm_get_function_metadata("CUBEVALUE", &present, &md), 0);
  EXPECT_EQ(md.availability, FM_FUNCTION_UNAVAILABLE_STUB);
}

TEST(FormulonCApiFunctionMetadata, AvailabilityDistinguishesNonStubSpecialCases) {
  int32_t present = 0;
  fm_function_metadata_t md = FM_STRUCT_INIT(fm_function_metadata_t);
  ASSERT_EQ(fm_get_function_metadata("FILTERXML", &present, &md), 0);
  EXPECT_EQ(md.availability, FM_FUNCTION_IMPLEMENTED);

  ASSERT_EQ(fm_get_function_metadata("INFO", &present, &md), 0);
  EXPECT_EQ(md.availability, FM_FUNCTION_ENVIRONMENT_BOUND);

  ASSERT_EQ(fm_get_function_metadata("SUM", &present, &md), 0);
  EXPECT_EQ(md.availability, FM_FUNCTION_IMPLEMENTED);
}

TEST(FormulonCApiFunctionMetadata, NullArgsReturnBindingNullPointer) {
  int32_t present = 0;
  fm_function_metadata_t md = FM_STRUCT_INIT(fm_function_metadata_t);
  EXPECT_EQ(fm_get_function_metadata(nullptr, &present, &md),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_get_function_metadata("SUM", &present, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_get_function_metadata("SUM", nullptr, &md),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiFunctionMetadata, UnknownProfileIsRejectedWithoutMutation) {
  const fm_status_t invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  for (const char* id : {"", "de_DE", "mac-365-xx_XX", "win-365-en-US"}) {
    const char* localized = "sentinel";
    EXPECT_EQ(fm_localize_function_name("SUM", id, &localized), invalid) << id;
    EXPECT_EQ(localized, nullptr);
    EXPECT_NE(std::strstr(fm_get_last_error_context(), "profile_id="), nullptr) << id;

    const char* canonical = "sentinel";
    EXPECT_EQ(fm_canonicalize_function_name("SUM", id, &canonical), invalid) << id;
    EXPECT_EQ(canonical, nullptr);
  }
}

TEST(FormulonCApiFunctionMetadata, NullProfileReturnsBindingNullPointer) {
  const fm_status_t null_ptr = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
  const char* out = nullptr;
  EXPECT_EQ(fm_localize_function_name("SUM", nullptr, &out), null_ptr);
  EXPECT_EQ(fm_canonicalize_function_name("SUM", nullptr, &out), null_ptr);
  EXPECT_EQ(fm_localize_function_name(nullptr, "mac-365-de_DE", &out), null_ptr);
  EXPECT_EQ(fm_canonicalize_function_name("SUM", "mac-365-de_DE", nullptr), null_ptr);
}

TEST(FormulonCApiFunctionMetadata, GermanAndFrenchShowLocalizedNames) {
  const char* out = nullptr;
  for (const char* id : {"mac-365-de_DE", "win-365-de_DE"}) {
    ASSERT_EQ(fm_localize_function_name("sum", id, &out), 0);
    EXPECT_STREQ(out, "SUMME");
    ASSERT_EQ(fm_canonicalize_function_name("SUMME", id, &out), 0);
    EXPECT_STREQ(out, "SUM");
    // The canonical spelling stays accepted as a fallback.
    ASSERT_EQ(fm_canonicalize_function_name("sum", id, &out), 0);
    EXPECT_STREQ(out, "SUM");
  }
  ASSERT_EQ(fm_localize_function_name("XLOOKUP", "mac-365-fr_FR", &out), 0);
  EXPECT_STREQ(out, "RECHERCHEX");
  ASSERT_EQ(fm_canonicalize_function_name("SOMME", "mac-365-fr_FR", &out), 0);
  EXPECT_STREQ(out, "SUM");
  ASSERT_EQ(fm_canonicalize_function_name("TRIM", "mac-365-fr_FR", &out), 0);
  EXPECT_STREQ(out, "MIRR");
  // A localized spelling of another locale is not a function here.
  EXPECT_NE(fm_canonicalize_function_name("SUMME", "mac-365-fr_FR", &out), 0);
}

TEST(FormulonCApiFunctionMetadata, OtherLocalesShowCanonicalNames) {
  const char* out = nullptr;
  for (const char* id : {"mac-365-ja_JP", "win-365-ja_JP", "mac-365-en_US", "win-365-en_US", "mac-365-zh_CN",
                         "mac-365-ko_KR", "mac-365-th_TH"}) {
    ASSERT_EQ(fm_localize_function_name("SUM", id, &out), 0) << id;
    EXPECT_STREQ(out, "SUM") << id;
    ASSERT_EQ(fm_canonicalize_function_name("sum", id, &out), 0) << id;
    EXPECT_STREQ(out, "SUM") << id;
  }
  for (const char* id : {"mac-365-en_US", "mac-365-zh_CN", "mac-365-ko_KR", "mac-365-th_TH"}) {
    ASSERT_EQ(fm_localize_function_name("DOLLAR", id, &out), 0) << id;
    EXPECT_STREQ(out, "DOLLAR") << id;
  }
}

// ja-JP renames only the three functions its name column lists.
TEST(FormulonCApiFunctionMetadata, JapaneseRenamesTheListedFunctions) {
  const char* out = nullptr;
  ASSERT_EQ(fm_localize_function_name("DOLLAR", "mac-365-ja_JP", &out), 0);
  EXPECT_STREQ(out, "YEN");
  ASSERT_EQ(fm_canonicalize_function_name("YEN", "win-365-ja_JP", &out), 0);
  EXPECT_STREQ(out, "DOLLAR");
  ASSERT_EQ(fm_localize_function_name("SUM", "mac-365-ja_JP", &out), 0);
  EXPECT_STREQ(out, "SUM");
}

TEST(FormulonCApiFunctionMetadata, FunctionCountIsPositive) {
  EXPECT_GT(count_of(fm_get_function_count), 100U);
}

TEST(FormulonCApiFunctionMetadata, FunctionNamesAreSortedAndComplete) {
  const std::size_t count = count_of(fm_get_function_count);
  ASSERT_GT(count, 0U);
  const char* prev = nullptr;
  for (std::size_t i = 0; i < count; ++i) {
    const char* name = nullptr;
    ASSERT_EQ(fm_get_function_name_at(i, &name), 0);
    ASSERT_NE(name, nullptr);
    if (prev != nullptr) {
      EXPECT_LT(std::strcmp(prev, name), 0) << "names not sorted at idx=" << i;
    }
    prev = name;
  }
}

TEST(FormulonCApiFunctionMetadata, NameAtOutOfRangeReturnsIndexOutOfRange) {
  const char* out = nullptr;
  fm_status_t rc = fm_get_function_name_at(static_cast<std::size_t>(-1), &out);
  EXPECT_EQ(rc, FM_STATUS_INDEX_OUT_OF_RANGE);
}

// Lazy-dispatch forms (XLOOKUP, SUMIFS, ...) and parser special forms (LET,
// LAMBDA) are recognised by the evaluator but are NOT in the eager registry.
// The catalog must still enumerate them and resolve their metadata.
TEST(FormulonCApiFunctionMetadata, EnumerationIncludesLazyAndSpecialForms) {
  std::unordered_set<std::string> names;
  const std::size_t count = count_of(fm_get_function_count);
  for (std::size_t i = 0; i < count; ++i) {
    const char* name = nullptr;
    ASSERT_EQ(fm_get_function_name_at(i, &name), 0);
    ASSERT_NE(name, nullptr);
    names.insert(name);
  }
  for (const char* expected : {"XLOOKUP", "SUMIFS", "IFERROR", "INDEX", "OFFSET", "INDIRECT", "SORT", "UNIQUE",
                               "FILTER", "LET", "LAMBDA", "VLOOKUP"}) {
    EXPECT_TRUE(names.count(expected) != 0) << expected << " missing from catalog enumeration";
  }
}

TEST(FormulonCApiFunctionMetadata, LazyAndSpecialFormsResolveMetadata) {
  int32_t present = 0;
  struct Expected {
    const char* name;
    std::uint32_t min_arity;
    std::uint32_t max_arity;
  };
  for (const Expected& fn :
       {Expected{"XLOOKUP", 3U, 6U}, Expected{"SUMIFS", 3U, 0xFFFFFFFFU}, Expected{"IFERROR", 2U, 2U},
        Expected{"INDEX", 2U, 4U}, Expected{"OFFSET", 3U, 5U}, Expected{"INDIRECT", 1U, 2U}, Expected{"SORT", 1U, 4U},
        Expected{"UNIQUE", 1U, 3U}, Expected{"FILTER", 2U, 3U}, Expected{"IF", 2U, 3U}, Expected{"LET", 3U, 253U},
        Expected{"LAMBDA", 1U, 254U}, Expected{"VLOOKUP", 3U, 4U}}) {
    fm_function_metadata_t md = FM_STRUCT_INIT(fm_function_metadata_t);
    ASSERT_EQ(fm_get_function_metadata(fn.name, &present, &md), 0) << fn.name;
    ASSERT_EQ(present, 1) << fn.name << " did not resolve";
    ASSERT_NE(md.canonical_name, nullptr);
    EXPECT_STREQ(md.canonical_name, fn.name);
    EXPECT_EQ(md.min_arity, fn.min_arity) << fn.name;
    EXPECT_EQ(md.max_arity, fn.max_arity) << fn.name;
    EXPECT_EQ(md.availability, FM_FUNCTION_IMPLEMENTED);
  }
}

TEST(FormulonCApiFunctionMetadata, LazyFormLookupIsCaseInsensitive) {
  int32_t present = 0;
  fm_function_metadata_t md = FM_STRUCT_INIT(fm_function_metadata_t);
  ASSERT_EQ(fm_get_function_metadata("xlookup", &present, &md), 0);
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
  int32_t present = 0;
  const std::size_t count = count_of(fm_get_function_count);
  ASSERT_GT(count, 0U);
  for (const char* profile : {"win-365-en_US", "mac-365-zh_CN"}) {
    for (std::size_t i = 0; i < count; ++i) {
      const char* name = nullptr;
      ASSERT_EQ(fm_get_function_name_at(i, &name), 0);
      ASSERT_NE(name, nullptr);

      // metadata
      fm_function_metadata_t md = FM_STRUCT_INIT(fm_function_metadata_t);
      EXPECT_EQ(fm_get_function_metadata(name, &present, &md), 0) << name;
      EXPECT_EQ(present, 1) << name << " has no metadata";
      // Display text is host-supplied, never engine-owned. Asserted for
      // the whole catalog rather than one sample: the eager and the
      // lazy / special-form branches populate the result separately, so
      // a single function only covers one of them.
      EXPECT_EQ(md.signature_template, nullptr) << name << " carries engine-side signature text";
      EXPECT_EQ(md.description, nullptr) << name << " carries engine-side description text";

      // canonicalize: an enumerated name is already canonical, so it maps to
      // itself.
      const char* canonical = nullptr;
      ASSERT_EQ(fm_canonicalize_function_name(name, profile, &canonical), 0) << name << " did not canonicalize";
      ASSERT_NE(canonical, nullptr);
      EXPECT_STREQ(canonical, name);

      // localize: with no alias table it falls through to the canonical name.
      const char* localized = nullptr;
      ASSERT_EQ(fm_localize_function_name(name, profile, &localized), 0) << name << " did not localize";
      ASSERT_NE(localized, nullptr);
      EXPECT_STREQ(localized, name);
    }
  }
}

TEST(FormulonCApiFunctionMetadata, LazyAndSpecialFormsCanonicalizeCaseInsensitively) {
  for (const char* fn : {"xlookup", "sumifs", "let", "lambda", "filter"}) {
    const char* canonical = nullptr;
    ASSERT_EQ(fm_canonicalize_function_name(fn, "win-365-en_US", &canonical), 0) << fn << " did not canonicalize";
    ASSERT_NE(canonical, nullptr);
    std::string upper(fn);
    for (char& c : upper) {
      c = static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
    }
    EXPECT_EQ(std::string(canonical), upper);
  }
}
