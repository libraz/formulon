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
  EXPECT_EQ(fm_localize_formula(formula, profile, nullptr, &out), 0) << profile;
  return out != nullptr ? out : "<null>";
}

std::string canonicalize(const char* formula, const char* profile) {
  const char* out = nullptr;
  EXPECT_EQ(fm_canonicalize_formula(formula, profile, nullptr, &out), 0) << profile;
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
  ASSERT_EQ(fm_localize_formula("=SUM(1.5,2)", "mac-365-de_DE", nullptr, &first), 0);
  const std::string copy = first;
  const char* second = nullptr;
  ASSERT_EQ(fm_canonicalize_formula(copy.c_str(), "mac-365-de_DE", nullptr, &second), 0);
  EXPECT_STREQ(second, "=SUM(1.5,2)");
}

TEST(FormulonCApiLocale, CanonicalizeReportsSyntaxErrorsInTheLocalizedText) {
  const char* out = nullptr;
  fm_parse_error_info_t error = FM_STRUCT_INIT(fm_parse_error_info_t);
  int32_t present = 0;

  const char* truncated = "=SUMME(1;";
  const fm_status_t rc = fm_canonicalize_formula(truncated, "mac-365-de_DE", nullptr, &out);
  EXPECT_GE(rc, 1000);
  EXPECT_LE(rc, 1014);
  ASSERT_EQ(fm_get_last_parse_error(&present, &error), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(error.end, 9U);
  EXPECT_EQ(error.line, 1U);
  EXPECT_EQ(error.column, 10U);
  EXPECT_EQ(error.end_cp, 9U);
  EXPECT_EQ(error.column_cp, 10U);

  // U+1F600 is two UTF-16 units and one code point of the typed text.
  const char* astral = "=\"\xF0\x9F\x98\x80\"&SUMME(1;";
  EXPECT_GE(fm_canonicalize_formula(astral, "mac-365-de_DE", nullptr, &out), 1000);
  present = 0;
  ASSERT_EQ(fm_get_last_parse_error(&present, &error), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(error.end, 14U);
  EXPECT_EQ(error.column, 15U);
  EXPECT_EQ(error.end_cp, 13U);
  EXPECT_EQ(error.column_cp, 14U);

  // The unterminated string sits at offset 23 of the typed text, not of the canonical rewrite.
  const char* unterminated = "=WENN(A1>0;SUMME(B1:B3;\"x);0)";
  EXPECT_GE(fm_canonicalize_formula(unterminated, "mac-365-de_DE", nullptr, &out), 1000);
  present = 0;
  ASSERT_EQ(fm_get_last_parse_error(&present, &error), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(error.start, 23U);
}

TEST(FormulonCApiLocale, LocalizeReportsSyntaxErrorsInTheCanonicalText) {
  const char* out = nullptr;
  fm_parse_error_info_t error = FM_STRUCT_INIT(fm_parse_error_info_t);
  int32_t present = 0;
  const fm_status_t rc = fm_localize_formula("=SUM(1,", "mac-365-de_DE", nullptr, &out);
  EXPECT_GE(rc, 1000);
  EXPECT_LE(rc, 1014);
  ASSERT_EQ(fm_get_last_parse_error(&present, &error), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(error.end, 7U);
  EXPECT_EQ(error.column, 8U);
}

TEST(FormulonCApiLocale, FormulaConversionRejectsBadArguments) {
  const char* out = nullptr;
  EXPECT_EQ(fm_localize_formula(nullptr, "mac-365-de_DE", nullptr, &out), kNullPtr);
  EXPECT_EQ(fm_localize_formula("=1", nullptr, nullptr, &out), kNullPtr);
  EXPECT_EQ(fm_localize_formula("=1", "mac-365-de_DE", nullptr, nullptr), kNullPtr);
  EXPECT_EQ(fm_canonicalize_formula("=1", nullptr, nullptr, &out), kNullPtr);
  EXPECT_EQ(fm_localize_formula("=1", "de_DE", nullptr, &out), kInvalid);
  EXPECT_STREQ(fm_get_last_error_context(), "profile_id=de_DE");
  EXPECT_EQ(fm_canonicalize_formula("=1", "mac-365-xx_XX", nullptr, &out), kInvalid);
  EXPECT_STREQ(fm_get_last_error_context(), "profile_id=mac-365-xx_XX");
}

constexpr uint32_t kLastRow = 1048575U;
constexpr uint32_t kLastCol = 16383U;

fm_reference_notation_spec_t Notation(int32_t r1c1, int32_t has_base = 0, uint32_t base_row = 0,
                                      uint32_t base_col = 0) {
  fm_reference_notation_spec_t spec = FM_STRUCT_INIT(fm_reference_notation_spec_t);
  spec.r1c1 = r1c1;
  spec.has_base = has_base;
  spec.base_row = base_row;
  spec.base_col = base_col;
  return spec;
}

std::string localize_with(const char* formula, const char* profile, const fm_reference_notation_spec_t& spec) {
  const char* out = nullptr;
  EXPECT_EQ(fm_localize_formula(formula, profile, &spec, &out), 0) << profile << ": " << fm_get_last_error_message();
  return out != nullptr ? out : "<null>";
}

std::string canonicalize_with(const char* formula, const char* profile, const fm_reference_notation_spec_t& spec) {
  const char* out = nullptr;
  EXPECT_EQ(fm_canonicalize_formula(formula, profile, &spec, &out), 0)
      << profile << ": " << fm_get_last_error_message();
  return out != nullptr ? out : "<null>";
}

// The failure is `FM_STATUS_PARSER_INVALID_REFERENCE` over [start, end) of the text passed.
void ExpectInvalidReference(fm_status_t rc, uint32_t start, uint32_t end) {
  EXPECT_EQ(rc, FM_STATUS_PARSER_INVALID_REFERENCE);
  fm_parse_error_info_t error = FM_STRUCT_INIT(fm_parse_error_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_get_last_parse_error(&present, &error), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(error.code, FM_PARSE_ERROR_CODE_INVALID_REFERENCE);
  EXPECT_EQ(error.start, start);
  EXPECT_EQ(error.end, end);
}

TEST(FormulonCApiLocale, OneNotationSpecServesParseLocalizeAndCanonicalize) {
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_create_default(&wb), 0);
  const char* out = nullptr;
  fm_parsed_reference_info_t parsed = FM_STRUCT_INIT(fm_parsed_reference_info_t);
  const auto expect_all = [&](const fm_reference_notation_spec_t& spec, fm_status_t want) {
    const char* text = spec.r1c1 == FM_TRISTATE_TRUE ? "R[1]C[1]" : "A1";
    EXPECT_EQ(fm_workbook_parse_reference(wb, text, &spec, &parsed), want);
    EXPECT_EQ(fm_localize_formula("=1", "mac-365-de_DE", &spec, &out), want);
    EXPECT_EQ(fm_canonicalize_formula("=1", "mac-365-de_DE", &spec, &out), want);
  };
  expect_all(Notation(FM_TRISTATE_TRUE, 1, 4, 4), 0);
  expect_all(Notation(FM_TRISTATE_DEFAULT, 2), kInvalid);
  expect_all(Notation(FM_TRISTATE_DEFAULT, 1, kLastRow + 1U, 0), kInvalid);
  expect_all(Notation(FM_TRISTATE_TRUE, 1, 0, kLastCol + 1U), kInvalid);
  fm_reference_notation_spec_t small = Notation(FM_TRISTATE_DEFAULT);
  small.struct_size = 4U;
  expect_all(small, kInvalid);
  fm_workbook_destroy(wb);
}

TEST(FormulonCApiLocale, NotationSpecIsValidatedLikeParseReference) {
  const char* de = "mac-365-de_DE";
  const char* out = nullptr;
  for (const bool localize_side : {true, false}) {
    const auto call = [&](const fm_reference_notation_spec_t& spec) {
      return localize_side ? fm_localize_formula("=1", de, &spec, &out)
                           : fm_canonicalize_formula("=1", de, &spec, &out);
    };
    fm_reference_notation_spec_t small = Notation(FM_TRISTATE_DEFAULT);
    small.struct_size = 4U;
    EXPECT_EQ(call(small), kInvalid) << localize_side;
    EXPECT_EQ(call(Notation(3)), kInvalid) << localize_side;
    EXPECT_EQ(call(Notation(-1)), kInvalid) << localize_side;
    EXPECT_STREQ(fm_get_last_error_context(), "r1c1=-1 min=0 max=2");
    EXPECT_EQ(call(Notation(FM_TRISTATE_DEFAULT, 2)), kInvalid) << localize_side;
    // The base is checked when A1 is used too.
    EXPECT_EQ(call(Notation(FM_TRISTATE_DEFAULT, 1, kLastRow + 1U, 0)), kInvalid) << localize_side;
    EXPECT_EQ(call(Notation(FM_TRISTATE_TRUE, 1, 0, kLastCol + 1U)), kInvalid) << localize_side;
    // Without `has_base` the base fields are not read.
    EXPECT_EQ(call(Notation(FM_TRISTATE_TRUE, 0, 0xFFFFFFFFU, 0xFFFFFFFFU)), 0) << localize_side;
    EXPECT_EQ(call(Notation(FM_TRISTATE_TRUE, 1, kLastRow, kLastCol)), 0) << localize_side;
  }
  // The profile is checked before the record.
  fm_reference_notation_spec_t small = Notation(FM_TRISTATE_DEFAULT);
  small.struct_size = 0U;
  EXPECT_EQ(fm_localize_formula("=1", "nope", &small, &out), kInvalid);
  EXPECT_STREQ(fm_get_last_error_context(), "profile_id=nope");
}

TEST(FormulonCApiLocale, DefaultAndFalseNotationKeepA1) {
  const char* de = "mac-365-de_DE";
  const char* a1 = "=SUM(C3,$B$2,Sheet2!A:A)";
  const std::string shown = localize(a1, de);
  EXPECT_EQ(shown, "=SUMME(C3;$B$2;Sheet2!A:A)");
  for (const int32_t mode : {static_cast<int32_t>(FM_TRISTATE_DEFAULT), static_cast<int32_t>(FM_TRISTATE_FALSE)}) {
    EXPECT_EQ(localize_with(a1, de, Notation(mode)), shown) << mode;
    EXPECT_EQ(localize_with(a1, de, Notation(mode, 1, 4, 4)), shown) << mode;
    EXPECT_EQ(canonicalize_with(shown.c_str(), de, Notation(mode, 1, 4, 4)), a1) << mode;
  }
}

TEST(FormulonCApiLocale, R1C1RoundTripsWithABase) {
  struct Case {
    const char* profile;
    const char* a1;
    const char* local;
  };
  // Mac Excel's FormulaR1C1Local of formulas entered in E5.
  const Case cases[] = {
      {"mac-365-en_US", "=SUM(C3,$B$2)", "=SUM(R[-2]C[-2],R2C2)"},
      {"mac-365-de_DE", "=SUM(C3,$B$2)", "=SUMME(Z(-2)S(-2);Z2S2)"},
      {"mac-365-es_ES", "=SUM(C3,$B$2)", "=SUMA(F[-2]C[-2];F2C2)"},
      {"mac-365-de_DE", "=IF($A$1>0,$B$1,\"x;y\")", "=WENN(Z1S1>0;Z1S2;\"x;y\")"},
      {"mac-365-es_ES", "=IF($A$1>0,$B$1,\"x;y\")", "=SI(F1C1>0;F1C2;\"x;y\")"},
  };
  const fm_reference_notation_spec_t at_e5 = Notation(FM_TRISTATE_TRUE, 1, 4, 4);
  for (const Case& c : cases) {
    EXPECT_EQ(localize_with(c.a1, c.profile, at_e5), c.local) << c.profile;
    EXPECT_EQ(canonicalize_with(c.local, c.profile, at_e5), c.a1) << c.profile;
  }
  // Absolute references need no base.
  const fm_reference_notation_spec_t no_base = Notation(FM_TRISTATE_TRUE);
  EXPECT_EQ(localize_with("=$A$1+SUM($1:$2)", "mac-365-de_DE", no_base), "=Z1S1+SUMME(Z1:Z2)");
  EXPECT_EQ(canonicalize_with("=Z1S1+SUMME(Z1:Z2)", "mac-365-de_DE", no_base), "=$A$1+SUM($1:$2)");
}

TEST(FormulonCApiLocale, R1C1FailuresPointIntoTheTextPassed) {
  const char* out = nullptr;
  const fm_reference_notation_spec_t no_base = Notation(FM_TRISTATE_TRUE);
  ExpectInvalidReference(fm_localize_formula("=SUM(A1,$B$2)", "mac-365-en_US", &no_base, &out), 5U, 7U);
  ExpectInvalidReference(fm_canonicalize_formula("=1+Z(-1)S", "mac-365-de_DE", &no_base, &out), 3U, 9U);
  const fm_reference_notation_spec_t at_e5 = Notation(FM_TRISTATE_TRUE, 1, 4, 4);
  ExpectInvalidReference(fm_canonicalize_formula("=SUMME(Z(x)S;1)", "mac-365-de_DE", &at_e5, &out), 7U, 11U);

  // A syntax error is still a parser status placed in the text passed.
  const fm_status_t rc = fm_localize_formula("=SUM(1,", "mac-365-de_DE", &at_e5, &out);
  EXPECT_GE(rc, 1000);
  EXPECT_LE(rc, 1014);
  fm_parse_error_info_t error = FM_STRUCT_INIT(fm_parse_error_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_get_last_parse_error(&present, &error), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(error.end, 7U);
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
