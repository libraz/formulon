#include "eval/formula_localize.h"

#include <cstdint>
#include <filesystem>
#include <fstream>
#include <set>
#include <sstream>
#include <string>
#include <string_view>

#include "c_api/formulon_c.h"
#include "excel_locale.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "parser/reference.h"
#include "parser/token.h"
#include "parser/tokenizer.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace eval {
namespace {

std::string localize_in(ExcelLocale locale, std::string_view formula) {
  return localize_formula_text(formula, ExcelProfile{ExcelHost::kMac365, locale});
}

std::string canonicalize_in(ExcelLocale locale, std::string_view formula) {
  return canonicalize_formula_text(formula, ExcelProfile{ExcelHost::kMac365, locale});
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
  int32_t present = 0;
  const char* out = nullptr;
  ASSERT_EQ(fm_localize_function_name("sum", "mac-365-de_DE", &out), 0);
  EXPECT_STREQ(out, "SUMME");
  ASSERT_EQ(fm_localize_function_name("ABS", "mac-365-fr_FR", &out), 0);
  EXPECT_STREQ(out, "ABS");
  ASSERT_EQ(fm_canonicalize_function_name("SOMME", "mac-365-fr_FR", &out), 0);
  EXPECT_STREQ(out, "SUM");
  ASSERT_EQ(fm_canonicalize_function_name("TRIM", "mac-365-fr_FR", &out), 0);
  EXPECT_STREQ(out, "MIRR");
  // The canonical spelling stays accepted as a fallback.
  ASSERT_EQ(fm_canonicalize_function_name("sum", "mac-365-de_DE", &out), 0);
  EXPECT_STREQ(out, "SUM");
  ASSERT_EQ(fm_localize_function_name("DOLLAR", "mac-365-ja_JP", &out), 0);
  EXPECT_STREQ(out, "YEN");
  fm_function_metadata_info_t md = FM_STRUCT_INIT(fm_function_metadata_info_t);
  EXPECT_EQ(fm_get_function_metadata("SUM", &present, &md), 0);
  EXPECT_EQ(present, 1);
  EXPECT_NE(fm_canonicalize_function_name("SUMME", "mac-365-fr_FR", &out), 0);
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

TEST(FormulaLocalize, CanonicalizeInvertsTheCaptures) {
  EXPECT_EQ(canonicalize_in(ExcelLocale::kDeDE, "=SUMME(1,5;2)"), "=SUM(1.5,2)");
  EXPECT_EQ(canonicalize_in(ExcelLocale::kDeDE, "=SUMME({1,5.2,5;3,4.4})"), "=SUM({1.5,2.5;3.4,4})");
  EXPECT_EQ(canonicalize_in(ExcelLocale::kDeDE, "=WENN(WAHR;1;FALSCH)"), "=IF(TRUE,1,FALSE)");
  EXPECT_EQ(canonicalize_in(ExcelLocale::kDeDE, "=WENNFEHLER(#NV;1)"), "=IFERROR(#N/A,1)");
  EXPECT_EQ(canonicalize_in(ExcelLocale::kDeDE, "=DM(1)+0"), "=DOLLAR(1)+0");
  EXPECT_EQ(canonicalize_in(ExcelLocale::kFrFR, "=SI(VRAI;1,5;0)"), "=IF(TRUE,1.5,0)");
  EXPECT_EQ(canonicalize_in(ExcelLocale::kFrFR, "=ARRONDI(SOMME(A1:A3);2)"), "=ROUND(SUM(A1:A3),2)");
  // Text the rewrite does not own stays as written, with or without `=`.
  EXPECT_EQ(canonicalize_in(ExcelLocale::kDeDE, "SUMME(\"a;b\";'a;b'!A1)"), "SUM(\"a;b\",'a;b'!A1)");
  EXPECT_EQ(canonicalize_in(ExcelLocale::kDeDE, "=SUMME(T[[#Headers],[a;b]];1)"), "=SUM(T[[#Headers],[a;b]],1)");
  EXPECT_EQ(canonicalize_in(ExcelLocale::kDeDE, "= summe( 1 ; LOG10(A1) )"), "= SUM( 1 , LOG10(A1) )");
  EXPECT_EQ(canonicalize_in(ExcelLocale::kJaJP, "=IF(TRUE,1.5,{1,2;3,4})"), "=IF(TRUE,1.5,{1,2;3,4})");
}

TEST(FormulaLocalize, QualifiedCallableNamesRemainVerbatim) {
  const char* const qualifiers[] = {"Sheet1!", "'My Sheet'!", "[0]!", "[Book.xlsx]Sheet1!"};
  for (const char* qualifier : qualifiers) {
    const std::string canonical = std::string("=") + qualifier + "SUM(1,2)";
    const std::string localized = std::string("=") + qualifier + "SUM(1;2)";
    EXPECT_EQ(localize_in(ExcelLocale::kDeDE, canonical), localized) << qualifier;

    const std::string localized_builtin = std::string("=") + qualifier + "SUMME(1;2)";
    const std::string canonicalized = std::string("=") + qualifier + "SUMME(1,2)";
    EXPECT_EQ(canonicalize_in(ExcelLocale::kDeDE, localized_builtin), canonicalized) << qualifier;
  }

  // French's localized spelling TRIM maps back to MIRR. A sheet-qualified
  // defined name must keep that spelling through both rewrite directions.
  EXPECT_EQ(localize_in(ExcelLocale::kFrFR, "=Sheet1!TRIM(1)"), "=Sheet1!TRIM(1)");
  EXPECT_EQ(canonicalize_in(ExcelLocale::kFrFR, "=Sheet1!TRIM(1)"), "=Sheet1!TRIM(1)");

  // TRUE/FALSE tokens after `!` are names in the qualifier's scope, not
  // localized Boolean literals, even when followed by a call parenthesis.
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=Sheet1!TRUE(1,2)"), "=Sheet1!TRUE(1;2)");
  EXPECT_EQ(canonicalize_in(ExcelLocale::kDeDE, "=Sheet1!WAHR(1;2)"), "=Sheet1!WAHR(1,2)");
}

TEST(FormulaLocalize, CanonicalizeRoundTripsLocalize) {
  const char* const formulas[] = {
      "=SUM(1.5,2)",       "=IF(TRUE,1.5,FALSE)",          "=SUM({1,2.5;3,4})", "=IFERROR(#N/A,\"x;y\")",
      "=SUM('a;b'!A1,B2)", "=SUM(T[[#Headers],[a]],0.25)", "SUM(1,2)",          "= XLOOKUP( 1 , A:A , B:B )"};
  for (const auto& entry : detail::kExcelProfileIds) {
    for (const char* formula : formulas) {
      const std::string shown = localize_formula_text(formula, entry.profile);
      EXPECT_EQ(canonicalize_formula_text(shown, entry.profile), formula) << entry.id << ": " << shown;
    }
  }
}

TEST(FormulaLocalize, KeepsUnlocalizedTextVerbatim) {
  // Whitespace and lowercase spellings survive; LOG10 lexes as a cell reference.
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "= sum( 1 , LOG10(A1) )"), "= SUMME( 1 ; LOG10(A1) )");
  // Structured-reference items take the locale's specifier words and list separator.
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=SUM(T[[#Headers],[A]],1)"), "=SUMME(T[[#Kopfzeilen];[A]];1)");
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

// `text` without its single quotes: the display requotes sheet qualifiers and
// names, which the entered text may quote differently.
bool is_bare_sheet_byte(char c) {
  const auto u = static_cast<unsigned char>(c);
  return u >= 0x80U || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '.' ||
         c == '_';
}

void append_requoted(std::string_view first, std::string_view last, const LocaleFacts& facts, std::string& out) {
  const bool quote = parser::local_sheet_needs_quoting_display(first, facts) ||
                     (!last.empty() && parser::local_sheet_needs_quoting_display(last, facts));
  std::string name(first);
  if (!last.empty()) {
    name.append(":").append(last);
  }
  if (!quote) {
    out.append(name);
    return;
  }
  out.push_back('\'');
  for (const char c : name) {
    out.push_back(c);
    if (c == '\'') {
      out.push_back('\'');
    }
  }
  out.push_back('\'');
}

// `formula` with each local sheet qualifier (`name!`, `'name'!`, `a:b!`,
// `'a:b'!`) requoted under the en-US display rule, found by a plain text scan
// independent of the tokenizer. Strings, `[...]` groups, error literals and
// partly quoted 3-D spans stay as written.
std::string with_display_quoting(std::string_view formula) {
  const LocaleFacts& facts = locale_facts(ExcelProfile{ExcelHost::kMac365, ExcelLocale::kEnUS});
  std::string out;
  std::size_t i = 0;
  int brackets = 0;
  while (i < formula.size()) {
    const char c = formula[i];
    if (c == '"') {
      const std::size_t start = i++;
      while (i < formula.size() && !(formula[i] == '"' && (i + 1 >= formula.size() || formula[i + 1] != '"'))) {
        i += formula[i] == '"' ? 2U : 1U;
      }
      out.append(formula.substr(start, ++i - start));
      continue;
    }
    if (c == '#') {  // an error literal (`#DIV/0!`, `#N/A`, `#NAME?`) or a spill `#`
      std::size_t j = i + 1;
      while (j < formula.size() && (is_bare_sheet_byte(formula[j]) || formula[j] == '/')) {
        ++j;
      }
      if (j < formula.size() && (formula[j] == '!' || formula[j] == '?')) {
        ++j;
      }
      out.append(formula.substr(i, j - i));
      i = j;
      continue;
    }
    if (c == '[' || c == ']') {
      brackets += c == '[' ? 1 : -1;
      out.push_back(c);
      ++i;
      continue;
    }
    if (brackets == 0 && c == '\'' && (i == 0 || formula[i - 1] != ']')) {
      std::string name;
      std::size_t j = i + 1;
      while (j < formula.size() && !(formula[j] == '\'' && (j + 1 >= formula.size() || formula[j + 1] != '\''))) {
        name.push_back(formula[j]);
        j += formula[j] == '\'' ? 2U : 1U;
      }
      if (j + 1 < formula.size() && formula[j + 1] == '!' && name.find('[') == std::string::npos) {
        const std::size_t colon = name.find(':');
        const std::string_view view(name);
        append_requoted(view.substr(0, colon), colon == std::string::npos ? std::string_view{} : view.substr(colon + 1),
                        facts, out);
        i = j + 1;
        continue;
      }
      out.append(formula.substr(i, j + 1 - i));
      i = j + 1;
      continue;
    }
    if (brackets == 0 && is_bare_sheet_byte(c) &&
        (i == 0 || (!is_bare_sheet_byte(formula[i - 1]) && formula[i - 1] != ']' && formula[i - 1] != '\''))) {
      std::size_t j = i;
      while (j < formula.size() && is_bare_sheet_byte(formula[j])) {
        ++j;
      }
      std::size_t k = j;
      if (k < formula.size() && formula[k] == ':' && k + 1 < formula.size() && is_bare_sheet_byte(formula[k + 1]) &&
          !parser::is_a1_cell_shaped(formula.substr(i, j - i))) {
        k += 1;
        while (k < formula.size() && is_bare_sheet_byte(formula[k])) {
          ++k;
        }
        if (k < formula.size() && formula[k] == '!') {
          append_requoted(formula.substr(i, j - i), formula.substr(j + 1, k - j - 1), facts, out);
          i = k;
          continue;
        }
      }
      if (j < formula.size() && formula[j] == '!') {
        append_requoted(formula.substr(i, j - i), {}, facts, out);
      } else {
        out.append(formula.substr(i, j - i));
      }
      i = j;
      continue;
    }
    out.push_back(c);
    ++i;
  }
  return out;
}

// With en-US facts the rewrite is the identity on every formula the oracle
// corpus enters, except exponent literals Excel redisplays in normalized form
// and sheet qualifiers, which take the en-US display quoting.
TEST(FormulaLocalize, EnglishRewriteIsIdentityOnGoldenFormulas) {
  const std::filesystem::path targets = std::filesystem::path(FORMULON_FIXTURES_DIR) / ".." / "oracle" / "targets";
  std::set<std::string> formulas;
  for (const auto& target : std::filesystem::directory_iterator(targets)) {
    const std::filesystem::path golden = target.path() / "golden";
    if (!std::filesystem::is_directory(golden)) {
      continue;
    }
    for (const auto& file : std::filesystem::directory_iterator(golden)) {
      // That suite enters deliberately non-canonical spellings to pin Excel's re-spelling.
      if (file.path().filename() == "entry_canonical.golden.json") {
        continue;
      }
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
    EXPECT_EQ(localize_in(ExcelLocale::kEnUS, formula), with_display_quoting(formula)) << formula;
    ++checked;
  }
  EXPECT_GT(checked, formulas.size() * 9 / 10);
}

struct SpecifierRow {
  ExcelLocale locale;
  const char* all;
  const char* data;
  const char* headers;
  const char* totals;
  char separator;
};

const SpecifierRow kSpecifierRows[] = {
    {ExcelLocale::kJaJP, "すべて", "データ", "見出し", "集計", ','},
    {ExcelLocale::kEnUS, "All", "Data", "Headers", "Totals", ','},
    {ExcelLocale::kDeDE, "Alle", "Daten", "Kopfzeilen", "Ergebnisse", ';'},
    {ExcelLocale::kFrFR, "Tout", "Données", "En-têtes", "Totaux", ';'},
    {ExcelLocale::kEsES, "Todo", "Datos", "Encabezados", "Totales", ';'},
    {ExcelLocale::kEsMX, "Todo", "Datos", "Encabezados", "Totales", ','},
    {ExcelLocale::kItIT, "Tutti", "Dati", "Intestazioni", "Totali", ';'},
    {ExcelLocale::kPtBR, "Tudo", "Dados", "Cabeçalhos", "Totais", ';'},
    {ExcelLocale::kRuRU, "Все", "Данные", "Заголовки", "Итоги", ';'},
    {ExcelLocale::kNlNL, "Alles", "Gegevens", "Kopteksten", "Totalen", ';'},
    {ExcelLocale::kZhCN, "全部", "数据", "标题", "汇总", ','},
    {ExcelLocale::kZhTW, "全部", "資料", "標題", "總計", ','},
    {ExcelLocale::kKoKR, "모두", "데이터", "머리글", "요약", ','},
    {ExcelLocale::kThTH, "All", "Data", "Headers", "Totals", ','},
};

TEST(FormulaLocalize, StructuredRefSpecifierWordsAndSeparatorInEveryLocale) {
  for (const SpecifierRow& row : kSpecifierRows) {
    const std::string sep(1, row.separator);
    const std::string cases[][2] = {
        {"=tblSales[[#All],[Qty]]", std::string("=tblSales[[#") + row.all + "]" + sep + "[Qty]]"},
        {"=tblSales[[#Data],[Qty]]", std::string("=tblSales[[#") + row.data + "]" + sep + "[Qty]]"},
        {"=tblSales[[#Headers],[Qty]]", std::string("=tblSales[[#") + row.headers + "]" + sep + "[Qty]]"},
        {"=tblSales[[#Totals],[Qty]]", std::string("=tblSales[[#") + row.totals + "]" + sep + "[Qty]]"},
        {"=tblSales[#Headers]", std::string("=tblSales[#") + row.headers + "]"},
        {"=tblSales[[#Headers],[#Data],[Qty]]",
         std::string("=tblSales[[#") + row.headers + "]" + sep + "[#" + row.data + "]" + sep + "[Qty]]"},
    };
    for (const auto& c : cases) {
      EXPECT_EQ(localize_in(row.locale, c[0]), c[1]) << c[0];
      EXPECT_EQ(canonicalize_in(row.locale, c[1]), c[0]) << c[1];
    }
  }
}

TEST(FormulaLocalize, StructuredRefSpecifierGermanTotals) {
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=tblSales[[#Totals],[Qty]]"), "=tblSales[[#Ergebnisse];[Qty]]");
  EXPECT_EQ(canonicalize_in(ExcelLocale::kDeDE, "=tblSales[[#Ergebnisse];[Qty]]"), "=tblSales[[#Totals],[Qty]]");
}

TEST(FormulaLocalize, StructuredRefSpecifierLeavesThisRowAndColumnNames) {
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=SUM(t[a,b],1.5)"), "=SUMME(t[a,b];1,5)");
  EXPECT_EQ(canonicalize_in(ExcelLocale::kDeDE, "=SUMME(t[a,b];1,5)"), "=SUM(t[a,b],1.5)");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=t[[a,b],[c]]"), "=t[[a,b];[c]]");
  EXPECT_EQ(localize_in(ExcelLocale::kDeDE, "=t[[#This Row],[a']b]]"), "=t[[#This Row];[a']b]]");
  EXPECT_EQ(localize_in(ExcelLocale::kJaJP, "=tblSales[@]"), "=tblSales[@]");
  EXPECT_EQ(localize_in(ExcelLocale::kJaJP, "=[@Qty]"), "=[@Qty]");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
