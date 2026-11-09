//
// C ABI locale surface: formula text conversion, locale facts, error names.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <iterator>
#include <string>

#include "c_api/formulon_c.h"
#include "formulon_c_count_of.h"
#include "gtest/gtest.h"
#include "utils/error.h"
#include "value.h"

namespace {

constexpr bool kWasm32 = sizeof(void*) == 4U;
static_assert(sizeof(fm_locale_facts_info_t) == (kWasm32 ? 60U : 104U), "fm_locale_facts_info_t ABI layout changed");

constexpr const char* kAllProfiles[] = {
    "mac-365-ja_JP", "win-365-ja_JP", "mac-365-en_US", "win-365-en_US", "mac-365-de_DE",
    "win-365-de_DE", "mac-365-fr_FR", "win-365-fr_FR", "mac-365-zh_CN", "win-365-zh_CN",
    "mac-365-ko_KR", "win-365-ko_KR", "mac-365-th_TH", "win-365-th_TH",
};

constexpr fm_status_t kInvalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNullPtr = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

std::string localize(const char* formula, const char* profile) {
  const char* out = nullptr;
  EXPECT_EQ(fm_localize_formula(formula, profile, &out), 0) << profile;
  return out != nullptr ? out : "<null>";
}

std::string canonicalize(const char* formula, const char* profile) {
  const char* out = nullptr;
  EXPECT_EQ(fm_canonicalize_formula(formula, profile, &out), 0) << profile;
  return out != nullptr ? out : "<null>";
}

TEST(FormulonCApiLocale, GermanFormulaConversion) {
  const char* de = "mac-365-de_DE";
  EXPECT_EQ(localize("=SUM(1.5,2)", de), "=SUMME(1,5;2)");
  EXPECT_EQ(canonicalize("=SUMME(1,5;2)", de), "=SUM(1.5,2)");
  EXPECT_EQ(localize("=SUM({1,2;3,4})", de), "=SUMME({1.2;3.4})");
  EXPECT_EQ(canonicalize("=SUMME({1.2;3.4})", de), "=SUM({1,2;3,4})");
  EXPECT_EQ(localize("=IF(TRUE,1,FALSE)", de), "=WENN(WAHR;1;FALSCH)");
  EXPECT_EQ(canonicalize("=WENN(WAHR;1;FALSCH)", de), "=IF(TRUE,1,FALSE)");
  EXPECT_EQ(localize("=IFERROR(#N/A,1)", de), "=WENNFEHLER(#NV;1)");
  EXPECT_EQ(canonicalize("=WENNFEHLER(#NV;1)", de), "=IFERROR(#N/A,1)");
}

TEST(FormulonCApiLocale, FrenchFormulaConversion) {
  const char* fr = "win-365-fr_FR";
  EXPECT_EQ(localize("=IF(TRUE,SUM(1.5,2),FALSE)", fr), "=SI(VRAI;SOMME(1,5;2);FAUX)");
  EXPECT_EQ(canonicalize("=SI(VRAI;SOMME(1,5;2);FAUX)", fr), "=IF(TRUE,SUM(1.5,2),FALSE)");
  EXPECT_EQ(localize("=SUM({1,2;3,4})", fr), "=SOMME({1.2;3.4})");
  EXPECT_EQ(canonicalize("=SOMME({1.2;3.4})", fr), "=SUM({1,2;3,4})");
}

TEST(FormulonCApiLocale, QuotedTextIsUntouchedAndStructuredItemsLocalize) {
  struct Case {
    const char* profile;
    const char* sum;
    const char* headers;
  };
  for (const Case& c : {Case{"mac-365-de_DE", "SUMME", "#Kopfzeilen"}, Case{"mac-365-fr_FR", "SOMME", "#En-têtes"}}) {
    const std::string shown = std::string("=") + c.sum + "(\"a,b;1.5\";'x,y;z'!A1)";
    EXPECT_EQ(localize("=SUM(\"a,b;1.5\",'x,y;z'!A1)", c.profile), shown) << c.profile;
    EXPECT_EQ(canonicalize(shown.c_str(), c.profile), "=SUM(\"a,b;1.5\",'x,y;z'!A1)") << c.profile;

    // Specifier words and the item separator follow the locale; a column name keeps its text.
    const std::string table = std::string("=") + c.sum + "(T[[" + c.headers + "];[a;b]];1)";
    EXPECT_EQ(localize("=SUM(T[[#Headers],[a;b]],1)", c.profile), table) << c.profile;
    EXPECT_EQ(canonicalize(table.c_str(), c.profile), "=SUM(T[[#Headers],[a;b]],1)") << c.profile;
  }
}

TEST(FormulonCApiLocale, EnglishAndJapaneseAreTheIdentity) {
  for (const char* profile : {"mac-365-en_US", "win-365-en_US", "mac-365-ja_JP", "win-365-ja_JP"}) {
    for (const char* formula : {"=SUM(1.5,2)", "=IF(TRUE,{1,2;3,4},#N/A)", "=SUM('a;b'!A1,\"x,y\")"}) {
      EXPECT_EQ(localize(formula, profile), formula) << profile;
      EXPECT_EQ(canonicalize(formula, profile), formula) << profile;
    }
  }
}

TEST(FormulonCApiLocale, EveryProfileRoundTrips) {
  const char* const formulas[] = {"=SUM(1.5,2)", "=IF(TRUE,{1,2.5;3,4},FALSE)", "=IFERROR(#N/A,\"x;y\")",
                                  "=XLOOKUP('a;b'!A1,A:A,B:B)", "=AVERAGE(T[[#Headers],[a]],0.25)"};
  for (const char* profile : kAllProfiles) {
    for (const char* formula : formulas) {
      const std::string shown = localize(formula, profile);
      EXPECT_EQ(canonicalize(shown.c_str(), profile), formula) << profile << ": " << shown;
    }
  }
}

TEST(FormulonCApiLocale, FormulaBufferIsReplacedByTheNextCall) {
  const char* first = nullptr;
  ASSERT_EQ(fm_localize_formula("=SUM(1.5,2)", "mac-365-de_DE", &first), 0);
  const std::string copy = first;
  const char* second = nullptr;
  ASSERT_EQ(fm_canonicalize_formula(copy.c_str(), "mac-365-de_DE", &second), 0);
  EXPECT_STREQ(second, "=SUM(1.5,2)");
}

TEST(FormulonCApiLocale, CanonicalizeReportsSyntaxErrorsInTheLocalizedText) {
  const char* out = nullptr;
  fm_parse_error_info_t error = FM_STRUCT_INIT(fm_parse_error_info_t);
  int32_t present = 0;

  const char* truncated = "=SUMME(1;";
  const fm_status_t rc = fm_canonicalize_formula(truncated, "mac-365-de_DE", &out);
  EXPECT_GE(rc, 1000);
  EXPECT_LE(rc, 1014);
  ASSERT_EQ(fm_get_last_parse_error(&present, &error), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(error.end, 9U);
  EXPECT_EQ(error.line, 1U);
  EXPECT_EQ(error.column, 10U);

  // The unterminated string sits at offset 23 of the typed text, not of the canonical rewrite.
  const char* unterminated = "=WENN(A1>0;SUMME(B1:B3;\"x);0)";
  EXPECT_GE(fm_canonicalize_formula(unterminated, "mac-365-de_DE", &out), 1000);
  present = 0;
  ASSERT_EQ(fm_get_last_parse_error(&present, &error), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(error.start, 23U);
}

TEST(FormulonCApiLocale, LocalizeReportsSyntaxErrorsInTheCanonicalText) {
  const char* out = nullptr;
  fm_parse_error_info_t error = FM_STRUCT_INIT(fm_parse_error_info_t);
  int32_t present = 0;
  const fm_status_t rc = fm_localize_formula("=SUM(1,", "mac-365-de_DE", &out);
  EXPECT_GE(rc, 1000);
  EXPECT_LE(rc, 1014);
  ASSERT_EQ(fm_get_last_parse_error(&present, &error), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(error.end, 7U);
  EXPECT_EQ(error.column, 8U);
}

TEST(FormulonCApiLocale, FormulaConversionRejectsBadArguments) {
  const char* out = nullptr;
  EXPECT_EQ(fm_localize_formula(nullptr, "mac-365-de_DE", &out), kNullPtr);
  EXPECT_EQ(fm_localize_formula("=1", nullptr, &out), kNullPtr);
  EXPECT_EQ(fm_localize_formula("=1", "mac-365-de_DE", nullptr), kNullPtr);
  EXPECT_EQ(fm_canonicalize_formula("=1", nullptr, &out), kNullPtr);
  EXPECT_EQ(fm_localize_formula("=1", "de_DE", &out), kInvalid);
  EXPECT_STREQ(fm_get_last_error_context(), "profile_id=de_DE");
  EXPECT_EQ(fm_canonicalize_formula("=1", "mac-365-xx_XX", &out), kInvalid);
  EXPECT_STREQ(fm_get_last_error_context(), "profile_id=mac-365-xx_XX");
}

TEST(FormulonCApiLocale, FactsForGermanFrenchAndJapanese) {
  fm_locale_facts_info_t de = FM_STRUCT_INIT(fm_locale_facts_info_t);
  ASSERT_EQ(fm_get_locale_facts("mac-365-de_DE", &de), 0);
  EXPECT_STREQ(de.decimal_separator, ",");
  EXPECT_STREQ(de.group_separator, ".");
  EXPECT_STREQ(de.list_separator, ";");
  EXPECT_STREQ(de.array_column_separator, ".");
  EXPECT_STREQ(de.array_row_separator, ";");
  EXPECT_STREQ(de.true_name, "WAHR");
  EXPECT_STREQ(de.false_name, "FALSCH");
  EXPECT_EQ(de.date_order, FM_DATE_ORDER_DMY);
  EXPECT_EQ(de.currency_suffix, 1);
  EXPECT_EQ(de.currency_space, 1);
  EXPECT_EQ(de.currency_default_decimals, 2);
  EXPECT_EQ(de.measured, 1);

  fm_locale_facts_info_t fr = FM_STRUCT_INIT(fm_locale_facts_info_t);
  ASSERT_EQ(fm_get_locale_facts("mac-365-fr_FR", &fr), 0);
  EXPECT_STREQ(fr.decimal_separator, ",");
  EXPECT_STREQ(fr.true_name, "VRAI");
  EXPECT_STREQ(fr.false_name, "FAUX");
  EXPECT_EQ(fr.date_order, FM_DATE_ORDER_DMY);

  fm_locale_facts_info_t ja = FM_STRUCT_INIT(fm_locale_facts_info_t);
  ASSERT_EQ(fm_get_locale_facts("mac-365-ja_JP", &ja), 0);
  EXPECT_STREQ(ja.decimal_separator, ".");
  EXPECT_STREQ(ja.group_separator, ",");
  EXPECT_STREQ(ja.list_separator, ",");
  EXPECT_STREQ(ja.array_column_separator, ",");
  EXPECT_STREQ(ja.array_row_separator, ";");
  EXPECT_STREQ(ja.true_name, "TRUE");
  EXPECT_EQ(ja.date_order, FM_DATE_ORDER_YMD);
  EXPECT_EQ(ja.currency_suffix, 0);
}

TEST(FormulonCApiLocale, MeasuredFlagFollowsTheProfile) {
  fm_locale_facts_info_t facts = FM_STRUCT_INIT(fm_locale_facts_info_t);
  ASSERT_EQ(fm_get_locale_facts("win-365-de_DE", &facts), 0);
  EXPECT_EQ(facts.measured, 0);
  ASSERT_EQ(fm_get_locale_facts("win-365-ja_JP", &facts), 0);
  EXPECT_EQ(facts.measured, 1);
  ASSERT_EQ(fm_get_locale_facts("win-365-en_US", &facts), 0);
  EXPECT_EQ(facts.measured, 0);
  for (const char* profile : kAllProfiles) {
    ASSERT_EQ(fm_get_locale_facts(profile, &facts), 0) << profile;
    const bool measured = std::strncmp(profile, "mac-", 4) == 0 || std::strcmp(profile, "win-365-ja_JP") == 0;
    EXPECT_EQ(facts.measured, measured ? 1 : 0) << profile;
  }
}

TEST(FormulonCApiLocale, FactsRejectBadArguments) {
  fm_locale_facts_info_t facts = FM_STRUCT_INIT(fm_locale_facts_info_t);
  EXPECT_EQ(fm_get_locale_facts(nullptr, &facts), kNullPtr);
  EXPECT_EQ(fm_get_locale_facts("mac-365-de_DE", nullptr), kNullPtr);
  EXPECT_EQ(fm_get_locale_facts("nope", &facts), kInvalid);
  EXPECT_STREQ(fm_get_last_error_context(), "profile_id=nope");
}

TEST(FormulonCApiLocale, ErrorNames) {
  EXPECT_EQ(count_of(fm_get_locale_error_name_count), std::size(formulon::kErrorTable));
  const size_t na = static_cast<size_t>(formulon::ErrorCode::NA);
  const size_t name = static_cast<size_t>(formulon::ErrorCode::Name);
  ASSERT_EQ(name, 4U);

  const char* canonical = nullptr;
  const char* localized = nullptr;
  int32_t measured = -1;
  ASSERT_EQ(fm_get_locale_error_name_at("mac-365-de_DE", na, &canonical, &localized, &measured), 0);
  EXPECT_STREQ(canonical, "#N/A");
  EXPECT_STREQ(localized, "#NV");
  EXPECT_EQ(measured, 1);

  ASSERT_EQ(fm_get_locale_error_name_at("win-365-de_DE", na, &canonical, &localized, &measured), 0);
  EXPECT_STREQ(localized, "#NV");
  EXPECT_EQ(measured, 0);

  ASSERT_EQ(fm_get_locale_error_name_at("mac-365-de_DE", name, &canonical, &localized, &measured), 0);
  EXPECT_STREQ(canonical, "#NAME?");
  EXPECT_EQ(measured, 1);

  ASSERT_EQ(fm_get_locale_error_name_at("mac-365-de_DE", static_cast<size_t>(formulon::ErrorCode::Field), &canonical,
                                        &localized, &measured),
            0);
  EXPECT_STREQ(canonical, "#FIELD!");
  EXPECT_EQ(measured, 0);

  ASSERT_EQ(fm_get_locale_error_name_at("mac-365-fr_FR", static_cast<size_t>(formulon::ErrorCode::Div0), &canonical,
                                        &localized, &measured),
            0);
  EXPECT_STREQ(localized, "#DIV/0!");
  EXPECT_EQ(measured, 1);

  EXPECT_EQ(fm_get_locale_error_name_at("mac-365-de_DE", count_of(fm_get_locale_error_name_count), &canonical,
                                        &localized, &measured),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_get_locale_error_name_at("bogus", na, &canonical, &localized, &measured), kInvalid);
  EXPECT_EQ(fm_get_locale_error_name_at(nullptr, na, &canonical, &localized, &measured), kNullPtr);
  EXPECT_EQ(fm_get_locale_error_name_at("mac-365-de_DE", na, &canonical, &localized, nullptr), kNullPtr);
}

TEST(FormulonCApiLocale, CollationMeasuredFollowsTheHost) {
  fm_locale_facts_info_t facts = FM_STRUCT_INIT(fm_locale_facts_info_t);
  ASSERT_EQ(fm_get_locale_facts("mac-365-ja_JP", &facts), 0);
  EXPECT_EQ(facts.collation_measured, 1);
  ASSERT_EQ(fm_get_locale_facts("win-365-ja_JP", &facts), 0);
  EXPECT_EQ(facts.measured, 1);
  EXPECT_EQ(facts.collation_measured, 0);
  ASSERT_EQ(fm_get_locale_facts("win-365-en_US", &facts), 0);
  EXPECT_EQ(facts.collation_measured, 0);
}

}  // namespace
