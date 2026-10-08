#include "eval/formula_localize.h"

#include <cstdint>
#include <filesystem>
#include <fstream>
#include <set>
#include <sstream>
#include <string>
#include <string_view>

#include "c_api/formulon_c.h"
#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "parser/token.h"
#include "parser/tokenizer.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace eval {
namespace {

std::string localize_in(ExcelLocale locale, std::string_view formula) {
  ScopedEvalProfile scope(ExcelProfile{ExcelHost::kMac365, locale});
  return localize_formula_text(formula);
}

TEST(FormulaLocalize, FunctionNameTable) {
  EXPECT_STREQ(localized_function_name(ExcelLocale::kDeDE, "SUM"), "SUMME");
  EXPECT_STREQ(localized_function_name(ExcelLocale::kFrFR, "xlookup"), "RECHERCHEX");
  EXPECT_STREQ(localized_function_name(ExcelLocale::kJaJP, "DOLLAR"), "YEN");
  EXPECT_STREQ(localized_function_name(ExcelLocale::kJaJP, "USDOLLAR"), "DOLLAR");
  EXPECT_EQ(localized_function_name(ExcelLocale::kJaJP, "SUM"), nullptr);
  EXPECT_EQ(localized_function_name(ExcelLocale::kEnUS, "SUM"), nullptr);
  EXPECT_EQ(localized_function_name(ExcelLocale::kDeDE, "ABS"), nullptr);

  EXPECT_STREQ(canonical_function_name(ExcelLocale::kDeDE, "summe"), "SUM");
  EXPECT_STREQ(canonical_function_name(ExcelLocale::kDeDE, "Z\xC3\x84HLENWENN"), "COUNTIF");
  // fr spells MIRR as TRIM; the localized spelling wins in that locale.
  EXPECT_STREQ(canonical_function_name(ExcelLocale::kFrFR, "TRIM"), "MIRR");
  EXPECT_EQ(canonical_function_name(ExcelLocale::kFrFR, "SUM"), nullptr);
  EXPECT_EQ(canonical_function_name(ExcelLocale::kZhCN, "SUM"), nullptr);
}

TEST(FormulaLocalize, CatalogApiReadsTheSameTable) {
  const char* out = nullptr;
  ASSERT_EQ(fm_function_localize("sum", FM_LOCALE_DE_DE, &out), 0);
  EXPECT_STREQ(out, "SUMME");
  ASSERT_EQ(fm_function_localize("ABS", FM_LOCALE_FR_FR, &out), 0);
  EXPECT_STREQ(out, "ABS");
  ASSERT_EQ(fm_function_canonicalize("SOMME", FM_LOCALE_FR_FR, &out), 0);
  EXPECT_STREQ(out, "SUM");
  ASSERT_EQ(fm_function_canonicalize("TRIM", FM_LOCALE_FR_FR, &out), 0);
  EXPECT_STREQ(out, "MIRR");
  // The canonical spelling stays accepted as a fallback.
  ASSERT_EQ(fm_function_canonicalize("sum", FM_LOCALE_DE_DE, &out), 0);
  EXPECT_STREQ(out, "SUM");
  ASSERT_EQ(fm_function_localize("DOLLAR", FM_LOCALE_JA_JP, &out), 0);
  EXPECT_STREQ(out, "DOLLAR");
  fm_function_metadata_t md{};
  EXPECT_EQ(fm_function_metadata("SUM", FM_LOCALE_FR_FR, &md), 0);
  EXPECT_NE(fm_function_canonicalize("SUMME", FM_LOCALE_FR_FR, &out), 0);
}

// Expected values are the locale_tokens.formulatext_* captures.
TEST(FormulaLocalize, GermanCaptures) {
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=1.5+2"), "=1,5+2");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=SUM({1,2;3,4})"), "=SUMME({1.2;3.4})");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=INDIRECT(\"R1C1\",FALSE)"), "=INDIREKT(\"R1C1\";FALSCH)");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=IF(TRUE,1,0)"), "=WENN(WAHR;1;0)");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=1.5E-3*2"), "=0,0015*2");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=IFERROR(#N/A,1)"), "=WENNFEHLER(#NV;1)");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=CONCAT(\"a,b\",\"1.5\")"), "=TEXTKETTE(\"a,b\";\"1.5\")");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=XLOOKUP(1,{1,2},{3,4})"), "=XVERWEIS(1;{1.2};{3.4})");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=SUM((C1:C2,D1:D2))"), "=SUMME((C1:C2;D1:D2))");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=DOLLAR(1)+0"), "=DM(1)+0");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=DBCS(\"a\")&USDOLLAR(1)"), "=DBCS(\"a\")&USDOLLAR(1)");
}

TEST(FormulaLocalize, FrenchAndJapaneseCaptures) {
  EXPECT_EQ(localize_in(ExcelLocale::kFrFR, "=IF(TRUE,1,0)"), "=SI(VRAI;1;0)");
  EXPECT_EQ(localize_in(ExcelLocale::kFrFR, "=IFERROR(#N/A,1)"), "=SIERREUR(#N/A;1)");
  EXPECT_EQ(localize_in(ExcelLocale::kFrFR, "=ROUND(SUM(A1:A3),2)"), "=ARRONDI(SOMME(A1:A3);2)");
  EXPECT_EQ(localize_in(ExcelLocale::kJaJP, "=DOLLAR(1)+0"), "=YEN(1)+0");
  EXPECT_EQ(localize_in(ExcelLocale::kJaJP, "=DBCS(\"a\")&USDOLLAR(1)"), "=JIS(\"a\")&DOLLAR(1)");
  EXPECT_EQ(localize_in(ExcelLocale::kJaJP, "=1.5E-3*2"), "=0.0015*2");
}

TEST(FormulaLocalize, KeepsUnlocalizedTextVerbatim) {
  // Whitespace and lowercase spellings survive; LOG10 lexes as a cell reference.
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "= sum( 1 , LOG10(A1) )"), "= SUMME( 1 ; LOG10(A1) )");
  // No capture covers structured-reference separators or item specifiers.
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=SUM(T[[#Headers],[A]],1)"), "=SUMME(T[[#Headers],[A]];1)");
  // A name that is not a call keeps its spelling.
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=SUM+1"), "=SUM+1");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "='a,b'!A1"), "='a,b'!A1");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=\"unterminated,"), "=\"unterminated,");
}

// Appends every JSON string value keyed "formula" in `text` to `out`.
void collect_formulas(const std::string& text, std::set<std::string>& out) {
  constexpr std::string_view kKey = "\"formula\":";
  for (std::size_t at = text.find(kKey); at != std::string::npos; at = text.find(kKey, at + 1)) {
    std::size_t i = text.find('"', at + kKey.size());
    if (i == std::string::npos) {
      return;
    }
    std::string value;
    for (++i; i < text.size() && text[i] != '"'; ++i) {
      if (text[i] != '\\') {
        value.push_back(text[i]);
        continue;
      }
      const char esc = text[++i];
      if (esc != 'u') {
        value.push_back(esc == 'n' ? '\n' : esc == 't' ? '\t' : esc == 'r' ? '\r' : esc);
        continue;
      }
      auto cp = static_cast<std::uint32_t>(std::stoul(text.substr(i + 1, 4), nullptr, 16));
      i += 4;
      if (cp >= 0xD800 && cp < 0xDC00 && text.compare(i + 1, 2, "\\u") == 0) {
        const auto low = static_cast<std::uint32_t>(std::stoul(text.substr(i + 3, 4), nullptr, 16));
        cp = 0x10000 + ((cp - 0xD800) << 10) + (low - 0xDC00);
        i += 6;
      }
      if (cp < 0x80) {
        value.push_back(static_cast<char>(cp));
      } else if (cp < 0x800) {
        value.push_back(static_cast<char>(0xC0 | (cp >> 6)));
        value.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
      } else if (cp < 0x10000) {
        value.push_back(static_cast<char>(0xE0 | (cp >> 12)));
        value.push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
        value.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
      } else {
        value.push_back(static_cast<char>(0xF0 | (cp >> 18)));
        value.push_back(static_cast<char>(0x80 | ((cp >> 12) & 0x3F)));
        value.push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
        value.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
      }
    }
    out.insert(std::move(value));
  }
}

// True when `formula` has a number literal written with an exponent, the one
// literal form the rewrite regenerates.
bool has_exponent_literal(const std::string& formula) {
  parser::Tokenizer tokenizer(formula);
  for (const parser::Token& token : tokenizer.tokens()) {
    if (token.kind == parser::TokenKind::Number && token.lexeme.find_first_of("Ee") != std::string_view::npos) {
      return true;
    }
  }
  return false;
}

// With en-US facts the rewrite is the identity on every formula the oracle
// corpus enters, except exponent literals Excel redisplays in normalized form.
TEST(FormulaLocalize, EnglishRewriteIsIdentityOnGoldenFormulas) {
  const std::filesystem::path targets = std::filesystem::path(FORMULON_FIXTURES_DIR) / ".." / "oracle" / "targets";
  std::set<std::string> formulas;
  for (const auto& target : std::filesystem::directory_iterator(targets)) {
    const std::filesystem::path golden = target.path() / "golden";
    if (!std::filesystem::is_directory(golden)) {
      continue;
    }
    for (const auto& file : std::filesystem::directory_iterator(golden)) {
      std::ifstream in(file.path());
      std::stringstream buffer;
      buffer << in.rdbuf();
      collect_formulas(buffer.str(), formulas);
    }
  }
  ASSERT_GT(formulas.size(), 1000U);
  std::size_t checked = 0;
  for (const std::string& formula : formulas) {
    if (has_exponent_literal(formula)) {
      continue;
    }
    EXPECT_EQ(localize_in(ExcelLocale::kEnUS, formula), formula);
    ++checked;
  }
  EXPECT_GT(checked, formulas.size() * 9 / 10);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
