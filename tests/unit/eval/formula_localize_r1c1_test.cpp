//
// Unit tests for the R1C1 notation of `localize_formula_text_r1c1` and
// `canonicalize_formula_text_r1c1`: the Mac Excel `FormulaR1C1Local` capture
// in every locale, the per-letter spellings, the error ranges, and the
// agreement with the AST printer `format_formula_r1c1`.

#include <cstdint>
#include <string>
#include <string_view>

#include "eval/builtin_names.h"
#include "eval/formula_localize.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_format_r1c1.h"
#include "parser/parse_error.h"
#include "parser/token.h"
#include "utils/arena.h"
#include "utils/formula_text.h"

namespace formulon {
namespace eval {
namespace {

struct CaptureRow {
  ExcelLocale locale;
  std::uint32_t row;
  std::uint32_t col;
  const char* a1;
  const char* local;
};

#include "data/formula_r1c1_capture.inc"

constexpr ExcelLocale kAllLocales[] = {
    ExcelLocale::kJaJP, ExcelLocale::kEnUS, ExcelLocale::kDeDE, ExcelLocale::kFrFR, ExcelLocale::kZhCN,
    ExcelLocale::kKoKR, ExcelLocale::kThTH, ExcelLocale::kEsES, ExcelLocale::kEsMX, ExcelLocale::kPtBR,
    ExcelLocale::kRuRU, ExcelLocale::kZhTW, ExcelLocale::kItIT, ExcelLocale::kNlNL,
};

ExcelProfile Mac(ExcelLocale locale) {
  return ExcelProfile{ExcelHost::kMac365, locale};
}

FormulaNotation At(std::uint32_t row, std::uint32_t col) {
  FormulaNotation notation;
  notation.r1c1 = true;
  notation.has_base = true;
  notation.base_row = row;
  notation.base_col = col;
  return notation;
}

FormulaNotation NoBase() {
  FormulaNotation notation;
  notation.r1c1 = true;
  return notation;
}

std::string Localize(ExcelLocale locale, std::string_view formula, const FormulaNotation& notation) {
  const Expected<MappedFormulaText, parser::ParseFailure> out =
      localize_formula_text_r1c1(formula, Mac(locale), notation);
  return out ? out.value().text : std::string("<failed>");
}

std::string Canonicalize(ExcelLocale locale, std::string_view formula, const FormulaNotation& notation) {
  const Expected<MappedFormulaText, parser::ParseFailure> out =
      canonicalize_formula_text_r1c1(formula, Mac(locale), notation);
  return out ? out.value().text : std::string("<failed>");
}

parser::TextRange Units(std::uint32_t start, std::uint32_t end) {
  parser::TextRange range;
  range.start = start;
  range.end = end;
  return range;
}

// Expects `out` to fail with `code` over UTF-16 units [start, end) on line 1.
void ExpectFailure(const Expected<MappedFormulaText, parser::ParseFailure>& out, parser::ParseErrorCode code,
                   std::uint32_t start, std::uint32_t end) {
  ASSERT_FALSE(out) << out.value().text;
  EXPECT_EQ(out.error().code, code);
  EXPECT_EQ(out.error().range.start, start);
  EXPECT_EQ(out.error().range.end, end);
  EXPECT_EQ(out.error().range.line, 1U);
  EXPECT_EQ(out.error().range.column, start + 1U);
}

// Rows from Mac Excel `FormulaR1C1` (tests/unit/parser/ast_format_r1c1_test.cpp):
// the normalized A1 spelling, without the leading `=`, and the host cell.
struct ProbeRow {
  const char* a1;
  std::uint32_t row;
  std::uint32_t col;
};

constexpr ProbeRow kProbeRows[] = {
    {"A1", 4, 2},
    {"$A$1", 4, 2},
    {"$A1", 4, 2},
    {"A$1", 4, 2},
    {"C5", 4, 2},
    {"D6", 4, 2},
    {"B4", 4, 2},
    {"A1:B2", 4, 2},
    {"$A$1:B2", 4, 2},
    {"A1:$B$2", 4, 2},
    {"SUM(A:A)", 4, 2},
    {"SUM($A:$A)", 4, 2},
    {"SUM(A:C)", 4, 2},
    {"SUM(D:E)", 4, 2},
    {"SUM($A:E)", 4, 2},
    {"SUM(1:1)", 4, 2},
    {"SUM($1:$1)", 4, 2},
    {"SUM(4:6)", 4, 2},
    {"SUM($4:6)", 4, 2},
    {"Data!A1", 4, 2},
    {"Data!$A$1", 4, 2},
    {"Data!A1:B2", 4, 2},
    {"'My Sheet'!A1", 4, 2},
    {"'It''s'!A1", 4, 2},
    {"'My Sheet'!$B$2:C3", 4, 2},
    {"Data!A:A", 4, 2},
    {"Data!1:1", 4, 2},
    {"Data:'My Sheet'!A1", 4, 2},
    {"SUM(Data:'My Sheet'!A1:B2)", 4, 2},
    {"Tbl[@Col1]", 4, 2},
    {"SUM(Tbl[@[Col1]:[Col2]])", 4, 2},
    {"SUM(Tbl[#Data])", 4, 2},
    {"Tbl[#Headers]", 4, 2},
    {"Tbl[[#Totals],[Col2]]", 4, 2},
    {"MyName", 4, 2},
    {"MyName+A1", 4, 2},
    {"SUM(MyRange)", 4, 2},
    {"Data!LocalName", 4, 2},
    {"\"A1\"", 4, 2},
    {"\"R1C1 and A1:B2\"&A1", 4, 2},
    {"IF(A1=\"B2\",1,0)", 4, 2},
    {"\"Sheet1!A1\"", 4, 2},
    {"\"He said \"\"A1\"\"\"", 4, 2},
    {"A1+B2*C3", 4, 2},
    {"SUM(A1,B2,$C$3)", 4, 2},
    {"INDEX(A1:C3,2,2)", 4, 2},
    {"A1:B2 B2:C3", 4, 2},
    {"(A1,B2)", 4, 2},
    {"A1:INDEX(B:B,3)", 4, 2},
    {"XFD1048576", 4, 2},
    {"A1048576", 4, 2},
    {"XFD1", 4, 2},
    {"A1#", 4, 2},
    {"A1:A3", 4, 2},
    {"LET(x,A1,x+B2)", 4, 2},
    {"LAMBDA(a,a+A1)(2)", 4, 2},
    {"B2", 0, 0},
    {"Z100", 0, 0},
    {"A1", 1048575, 16383},
    {"XFD1048575", 1048575, 16383},
    {"A1", 1, 1},
    {"SUM(A:A)", 1, 1},
    {"[1]Sheet1!A1", 4, 2},
    {"[1]Sheet1!$A$1:B2", 4, 2},
    {"'[1]It''s'!A1", 4, 2},
};

// ---------------------------------------------------------------------------
// Excel capture
// ---------------------------------------------------------------------------

TEST(FormulaLocalizeR1C1, MatchesExcelCapture) {
  for (const CaptureRow& row : kForwardRows) {
    SCOPED_TRACE(testing::Message() << static_cast<int>(row.locale) << " " << row.a1);
    EXPECT_EQ(Localize(row.locale, row.a1, At(row.row, row.col)), row.local);
  }
}

bool RefusesInvariantLetters(ExcelLocale locale) {
  switch (locale) {
    case ExcelLocale::kDeDE:
    case ExcelLocale::kFrFR:
    case ExcelLocale::kEsES:
    case ExcelLocale::kEsMX:
    case ExcelLocale::kNlNL:
    case ExcelLocale::kPtBR:
      return true;
    default:
      return false;
  }
}

// What Excel stores for the invariant row typed at (`row`, `col`) in en-US.
const char* EnglishReading(std::uint32_t row, std::uint32_t col) {
  for (const CaptureRow& en : kInvariantRows) {
    if (en.locale == ExcelLocale::kEnUS && en.row == row && en.col == col) {
      return en.a1;
    }
  }
  return "";
}

TEST(FormulaLocalizeR1C1, CanonicalizesExcelCapture) {
  for (const CaptureRow& row : kInverseRows) {
    SCOPED_TRACE(testing::Message() << static_cast<int>(row.locale) << " " << row.local);
    EXPECT_EQ(Canonicalize(row.locale, row.local, At(row.row, row.col)), row.a1);
  }
  // Invariant R1C1 text: what Excel accepted reads the same; in the locales
  // with their own letters, what it refused is never read as the invariant
  // references.
  for (const CaptureRow& row : kInvariantRows) {
    SCOPED_TRACE(testing::Message() << static_cast<int>(row.locale) << " " << row.local);
    const std::string canonical = Canonicalize(row.locale, row.local, At(row.row, row.col));
    if (*row.a1 != '\0') {
      EXPECT_EQ(canonical, row.a1);
    } else if (RefusesInvariantLetters(row.locale)) {
      EXPECT_NE(canonical, EnglishReading(row.row, row.col));
    }
  }
}

// ---------------------------------------------------------------------------
// Spelling per letter set
// ---------------------------------------------------------------------------

struct SpellingRow {
  const char* a1;
  const char* de;
  const char* fr;
  const char* es;
  const char* nl;
  const char* pt;
};

// Entered at (4, 2), i.e. C5.
constexpr SpellingRow kSpellingRows[] = {
    {"=A1", "=Z(-4)S(-2)", "=L(-4)C(-2)", "=F[-4]C[-2]", "=R[-4]K[-2]", "=L[-4]C[-2]"},
    {"=C5", "=ZS", "=LC", "=FC", "=RK", "=LC"},
    {"=D7", "=Z(2)S(1)", "=L(2)C(1)", "=F[2]C[1]", "=R[2]K[1]", "=L[2]C[1]"},
    {"=$A$1", "=Z1S1", "=L1C1", "=F1C1", "=R1K1", "=L1C1"},
    {"=$A1", "=Z(-4)S1", "=L(-4)C1", "=F[-4]C1", "=R[-4]K1", "=L[-4]C1"},
    {"=A$1", "=Z1S(-2)", "=L1C(-2)", "=F1C[-2]", "=R1K[-2]", "=L1C[-2]"},
    {"=A:A", "=S(-2)", "=C(-2)", "=C[-2]", "=K[-2]", "=C[-2]"},
    {"=$A:$A", "=S1", "=C1", "=C1", "=K1", "=C1"},
    {"=$A:E", "=S1:S(2)", "=C1:C(2)", "=C1:C[2]", "=K1:K[2]", "=C1:C[2]"},
    {"=C:C", "=S", "=C", "=C", "=K", "=C"},
    {"=5:5", "=Z", "=L", "=F", "=R", "=L"},
    {"=$1:$1", "=Z1", "=L1", "=F1", "=R1", "=L1"},
    {"=4:6", "=Z(-1):Z(1)", "=L(-1):L(1)", "=F[-1]:F[1]", "=R[-1]:R[1]", "=L[-1]:L[1]"},
    {"=A1:$B$2", "=Z(-4)S(-2):Z2S2", "=L(-4)C(-2):L2C2", "=F[-4]C[-2]:F2C2", "=R[-4]K[-2]:R2K2", "=L[-4]C[-2]:L2C2"},
};

TEST(FormulaLocalizeR1C1, SpellsEachLetterSet) {
  const struct {
    ExcelLocale locale;
    const char* SpellingRow::*column;
  } kColumns[] = {
      {ExcelLocale::kDeDE, &SpellingRow::de}, {ExcelLocale::kFrFR, &SpellingRow::fr},
      {ExcelLocale::kEsES, &SpellingRow::es}, {ExcelLocale::kNlNL, &SpellingRow::nl},
      {ExcelLocale::kPtBR, &SpellingRow::pt},
  };
  for (const SpellingRow& row : kSpellingRows) {
    for (const auto& column : kColumns) {
      SCOPED_TRACE(testing::Message() << static_cast<int>(column.locale) << " " << row.a1);
      const char* local = row.*column.column;
      EXPECT_EQ(Localize(column.locale, row.a1, At(4, 2)), local);
      EXPECT_EQ(Canonicalize(column.locale, local, At(4, 2)), row.a1);
    }
  }
}

TEST(FormulaLocalizeR1C1, FarOffsetsPrintUnwrappedAndReadWrapped) {
  EXPECT_EQ(Localize(ExcelLocale::kEnUS, "=XFD1048576", At(4, 2)), "=R[1048571]C[16381]");
  EXPECT_EQ(Localize(ExcelLocale::kEnUS, "=A1", At(1048575, 16383)), "=R[-1048575]C[-16383]");
  EXPECT_EQ(Canonicalize(ExcelLocale::kEnUS, "=R[1048571]C[16381]", At(4, 2)), "=XFD1048576");
  EXPECT_EQ(Canonicalize(ExcelLocale::kEnUS, "=R[-1048575]C[-16383]", At(1048575, 16383)), "=A1");
  // A short offset past the grid edge wraps on read.
  EXPECT_EQ(Canonicalize(ExcelLocale::kEnUS, "=R[-5]C[-3]", At(4, 2)), "=XFD1048576");
  EXPECT_EQ(Canonicalize(ExcelLocale::kDeDE, "=Z(1)S(1)", At(1048575, 16383)), "=A1");
}

TEST(FormulaLocalizeR1C1, RuArrayRoundTrips) {
  const std::string_view a1 = "=SUM($A$1,{1,2;3,4})";
  const std::string local = Localize(ExcelLocale::kRuRU, a1, At(4, 4));
  EXPECT_EQ(local, "=СУММ(R1C1;{1;2:3;4})");
  EXPECT_EQ(Canonicalize(ExcelLocale::kRuRU, local, At(4, 4)), a1);
  EXPECT_EQ(Canonicalize(ExcelLocale::kRuRU, "=СУММ(R[-1]C;{1:2})", At(4, 4)), "=SUM(E4,{1;2})");
}

TEST(FormulaLocalizeR1C1, TrimReferencesKeepTheirOperator) {
  const struct {
    const char* a1;
    const char* en;
    const char* de;
  } kRows[] = {
      {"=SUM(A1:.B2)", "=SUM(R[-4]C[-2]:.R[-3]C[-1])", "=SUMME(Z(-4)S(-2):.Z(-3)S(-1))"},
      {"=SUM(A1.:B2)", "=SUM(R[-4]C[-2].:R[-3]C[-1])", "=SUMME(Z(-4)S(-2).:Z(-3)S(-1))"},
      {"=SUM(A1.:.B2)", "=SUM(R[-4]C[-2].:.R[-3]C[-1])", "=SUMME(Z(-4)S(-2).:.Z(-3)S(-1))"},
      {"=SUM(A:.A)", "=SUM(C[-2]:.C[-2])", "=SUMME(S(-2):.S(-2))"},
  };
  for (const auto& row : kRows) {
    SCOPED_TRACE(row.a1);
    EXPECT_EQ(Localize(ExcelLocale::kEnUS, row.a1, At(4, 2)), row.en);
    EXPECT_EQ(Canonicalize(ExcelLocale::kEnUS, row.en, At(4, 2)), row.a1);
    EXPECT_EQ(Localize(ExcelLocale::kDeDE, row.a1, At(4, 2)), row.de);
    EXPECT_EQ(Canonicalize(ExcelLocale::kDeDE, row.de, At(4, 2)), row.a1);
  }
}

TEST(FormulaLocalizeR1C1, ReferenceShapes) {
  const struct {
    const char* a1;
    const char* local;
  } kRows[] = {
      {"=A1#", "=R[-4]C[-2]#"},
      {"=A1:INDEX(B:B,3)", "=R[-4]C[-2]:INDEX(C[-1],3)"},
      {"=A1:B2 B2:C3", "=R[-4]C[-2]:R[-3]C[-1] R[-3]C[-1]:R[-2]C"},
      {"=(A1,B2)", "=(R[-4]C[-2],R[-3]C[-1])"},
      {"=SUM(Sheet2:Sheet3!A1:B2)", "=SUM(Sheet2:Sheet3!R[-4]C[-2]:R[-3]C[-1])"},
      {"=Sheet2:Sheet3!$A:$B", "=Sheet2:Sheet3!C1:C2"},
      {"=[1]Sheet1!$A$1:B2", "=[1]Sheet1!R1C1:R[-3]C[-1]"},
      {"=[1]Sheet1!A:A", "=[1]Sheet1!C[-2]"},
      {"=Data!1:1", "=Data!R[-4]"},
      {"=SUM(Tbl[[#Totals],[Col2]])&\"A1\"", "=SUM(Tbl[[#Totals],[Col2]])&\"A1\""},
      {"=LET(x,A1,x+B2)", "=LET(x,R[-4]C[-2],x+R[-3]C[-1])"},
      {"=INDIRECT(\"A1\")+A1", "=INDIRECT(\"A1\")+R[-4]C[-2]"},
      {"= SUM( A1 , B2 )", "= SUM( R[-4]C[-2] , R[-3]C[-1] )"},
      {"=#REF!A1+1", "=#REF!R[-4]C[-2]+1"},
  };
  for (const auto& row : kRows) {
    SCOPED_TRACE(row.a1);
    EXPECT_EQ(Localize(ExcelLocale::kEnUS, row.a1, At(4, 2)), row.local);
    EXPECT_EQ(Canonicalize(ExcelLocale::kEnUS, row.local, At(4, 2)), row.a1);
  }
  // Structured references keep their brackets in a `( )` locale and take the
  // localized specifiers and separator.
  EXPECT_EQ(Localize(ExcelLocale::kDeDE, "=SUM(Tbl[[#Data],[Col2]],A1)&\"A1\"", At(4, 2)),
            "=SUMME(Tbl[[#Daten];[Col2]];Z(-4)S(-2))&\"A1\"");
  EXPECT_EQ(Canonicalize(ExcelLocale::kDeDE, "=SUMME(Tbl[[#Daten];[Col2]];Z(-4)S(-2))&\"A1\"", At(4, 2)),
            "=SUM(Tbl[[#Data],[Col2]],A1)&\"A1\"");
}

TEST(FormulaLocalizeR1C1, NotationOffKeepsTheA1Rewrite) {
  const std::string_view a1 = "=IF(A1>0,SUM(B1:B3,1.5),TRUE)";
  const ExcelProfile de = Mac(ExcelLocale::kDeDE);
  const Expected<MappedFormulaText, parser::ParseFailure> local = localize_formula_text_r1c1(a1, de, {});
  ASSERT_TRUE(local);
  EXPECT_EQ(local.value().text, localize_formula_text(a1, de));
  const Expected<MappedFormulaText, parser::ParseFailure> back =
      canonicalize_formula_text_r1c1(local.value().text, de, {});
  ASSERT_TRUE(back);
  EXPECT_EQ(back.value().text, a1);
}

// Empty text passes; a lone `=` does not parse, as in A1.
TEST(FormulaLocalizeR1C1, EmptyFormula) {
  EXPECT_EQ(Localize(ExcelLocale::kDeDE, "", NoBase()), "");
  EXPECT_EQ(Canonicalize(ExcelLocale::kDeDE, "", NoBase()), "");
  EXPECT_EQ(Localize(ExcelLocale::kDeDE, "=", NoBase()), "<failed>");
  EXPECT_EQ(Canonicalize(ExcelLocale::kDeDE, "=", NoBase()), "<failed>");
}

// ---------------------------------------------------------------------------
// Qualifiers, names and literals (Mac Excel de-DE / es-ES / nl-NL)
// ---------------------------------------------------------------------------

std::string LocalizeA1(ExcelLocale locale, std::string_view formula) {
  return localize_formula_text(formula, Mac(locale));
}

std::string CanonicalizeA1(ExcelLocale locale, std::string_view formula) {
  return canonicalize_formula_text(formula, Mac(locale));
}

TEST(FormulaLocalizeR1C1, RequotesSheetQualifiers) {
  const struct {
    const char* stored;
    const char* de_a1;
    const char* de_r1c1;
    const char* es_a1;
    const char* es_r1c1;
    const char* nl_a1;
    const char* nl_r1c1;
  } kRows[] = {
      {"='S2'!$A$1", "='S2'!$A$1", "='S2'!Z1S1", "='S2'!$A$1", "=S2!F1C1", "='S2'!$A$1", "=S2!R1K1"},
      {"=Z1S1!$A$1", "='Z1S1'!$A$1", "='Z1S1'!Z1S1", "=Z1S1!$A$1", "=Z1S1!F1C1", "=Z1S1!$A$1", "=Z1S1!R1K1"},
      {"=F1C1!$A$1", "=F1C1!$A$1", "=F1C1!Z1S1", "='F1C1'!$A$1", "='F1C1'!F1C1", "=F1C1!$A$1", "=F1C1!R1K1"},
      {"='R1K1'!$A$1", "='R1K1'!$A$1", "='R1K1'!Z1S1", "='R1K1'!$A$1", "='R1K1'!F1C1", "='R1K1'!$A$1", "='R1K1'!R1K1"},
      {"='R1C1'!$A$1", "='R1C1'!$A$1", "='R1C1'!Z1S1", "='R1C1'!$A$1", "='R1C1'!F1C1", "='R1C1'!$A$1", "='R1C1'!R1K1"},
      {"=SUM('Data:S2'!$A$1)", "=SUMME('Data:S2'!$A$1)", "=SUMME('Data:S2'!Z1S1)", "=SUMA('Data:S2'!$A$1)",
       "=SUMA(Data:S2!F1C1)", "=SOM('Data:S2'!$A$1)", "=SOM(Data:S2!R1K1)"},
      {"=SUM('S2:Z1S1'!$A$1)", "=SUMME('S2:Z1S1'!$A$1)", "=SUMME('S2:Z1S1'!Z1S1)", "=SUMA('S2:Z1S1'!$A$1)",
       "=SUMA(S2:Z1S1!F1C1)", "=SOM('S2:Z1S1'!$A$1)", "=SOM(S2:Z1S1!R1K1)"},
      {"=SUM('Data:My Sheet'!$A$1)", "=SUMME('Data:My Sheet'!$A$1)", "=SUMME('Data:My Sheet'!Z1S1)",
       "=SUMA('Data:My Sheet'!$A$1)", "=SUMA('Data:My Sheet'!F1C1)", "=SOM('Data:My Sheet'!$A$1)",
       "=SOM('Data:My Sheet'!R1K1)"},
  };
  for (const auto& row : kRows) {
    SCOPED_TRACE(row.stored);
    const struct {
      ExcelLocale locale;
      const char* a1;
      const char* r1c1;
    } kForms[] = {
        {ExcelLocale::kDeDE, row.de_a1, row.de_r1c1},
        {ExcelLocale::kEsES, row.es_a1, row.es_r1c1},
        {ExcelLocale::kNlNL, row.nl_a1, row.nl_r1c1},
    };
    for (const auto& form : kForms) {
      SCOPED_TRACE(static_cast<int>(form.locale));
      EXPECT_EQ(LocalizeA1(form.locale, row.stored), form.a1);
      EXPECT_EQ(Localize(form.locale, row.stored, NoBase()), form.r1c1);
      EXPECT_EQ(CanonicalizeA1(form.locale, form.a1), row.stored);
      EXPECT_EQ(Canonicalize(form.locale, form.r1c1, NoBase()), row.stored);
    }
  }
  // The invariant R1C1 display leaves `Data:S2` bare.
  EXPECT_EQ(Localize(ExcelLocale::kEnUS, "=SUM('Data:S2'!$A$1)", NoBase()), "=SUM(Data:S2!R1C1)");
  EXPECT_EQ(Canonicalize(ExcelLocale::kEnUS, "=SUM(Data:S2!R1C1)", NoBase()), "=SUM('Data:S2'!$A$1)");
}

TEST(FormulaLocalizeR1C1, NamesCollidingWithLocaleLetters) {
  // A defined name reading as a locale reference prints quoted in both forms.
  const struct {
    ExcelLocale locale;
    const char* stored;
    const char* local;
  } kQuoted[] = {
      {ExcelLocale::kDeDE, "=z+1", "='z'+1"},   {ExcelLocale::kDeDE, "=s+1", "='s'+1"},
      {ExcelLocale::kDeDE, "=zs+1", "='zs'+1"}, {ExcelLocale::kEsES, "=f+1", "='f'+1"},
      {ExcelLocale::kNlNL, "=k+1", "='k'+1"},   {ExcelLocale::kDeDE, "=f+1", "=f+1"},
      {ExcelLocale::kDeDE, "=r+1", "=r+1"},     {ExcelLocale::kEsES, "=z+1", "=z+1"},
  };
  for (const auto& row : kQuoted) {
    SCOPED_TRACE(row.stored);
    EXPECT_EQ(LocalizeA1(row.locale, row.stored), row.local);
    EXPECT_EQ(Localize(row.locale, row.stored, NoBase()), row.local);
  }
  // LET / LAMBDA names print as written.
  EXPECT_EQ(Localize(ExcelLocale::kDeDE, "=LET(z,1,z+1)", NoBase()), "=LET(z;1;z+1)");
  EXPECT_EQ(LocalizeA1(ExcelLocale::kDeDE, "=LET(z,1,z+1)"), "=LET(z;1;z+1)");
  EXPECT_EQ(Localize(ExcelLocale::kEsES, "=LET(f,1,f+1)", NoBase()), "=LET(f;1;f+1)");
  EXPECT_EQ(Localize(ExcelLocale::kDeDE, "=LET(s,LAMBDA(x,x+1),s(A1))", At(13, 3)),
            "=LET(s;LAMBDA(x;x+1);s(Z(-13)S(-3)))");

  // Every back-entry Excel refuses fails.
  const struct {
    ExcelLocale locale;
    const char* local;
  } kRefused[] = {
      {ExcelLocale::kDeDE, "=LET(z;1;z+1)"},
      {ExcelLocale::kDeDE, "=LET(s;1;s+1)"},
      {ExcelLocale::kDeDE, "=LET(zs;1;zs+1)"},
      {ExcelLocale::kEsES, "=LET(f;1;f+1)"},
      {ExcelLocale::kEsES, "=LET(c;1;c+1)"},
      {ExcelLocale::kEsES, "=LET(fc;1;fc+1)"},
      {ExcelLocale::kNlNL, "=LET(k;1;k+1)"},
      {ExcelLocale::kNlNL, "=LET(r;1;r+1)"},
      {ExcelLocale::kJaJP, "=LET(r,1,r+1)"},
      {ExcelLocale::kJaJP, "=LET(c,1,c+1)"},
      {ExcelLocale::kDeDE, "=LET(s;LAMBDA(x;x+1);s(2))"},
      {ExcelLocale::kDeDE, "=LET(s;LAMBDA(x;x+1);s(Z(-13)S(-3)))"},
      {ExcelLocale::kEsES, "=LET(f;LAMBDA(x;x+1);f(F[-14]C[-3]))"},
      {ExcelLocale::kDeDE, "='z'+1"},
      {ExcelLocale::kEsES, "=SUMA(F[Col])"},
  };
  for (const auto& row : kRefused) {
    SCOPED_TRACE(row.local);
    EXPECT_FALSE(canonicalize_formula_text_r1c1(row.local, Mac(row.locale), At(14, 3)));
  }
  // Letters that are no locale token round-trip, the invariant ones included.
  const struct {
    ExcelLocale locale;
    const char* local;
    const char* stored;
  } kAccepted[] = {
      {ExcelLocale::kDeDE, "=LET(r;1;r+1)", "=LET(r,1,r+1)"}, {ExcelLocale::kDeDE, "=LET(c;1;c+1)", "=LET(c,1,c+1)"},
      {ExcelLocale::kDeDE, "=LET(f;1;f+1)", "=LET(f,1,f+1)"}, {ExcelLocale::kEsES, "=LET(r;1;r+1)", "=LET(r,1,r+1)"},
      {ExcelLocale::kEsES, "=LET(s;1;s+1)", "=LET(s,1,s+1)"}, {ExcelLocale::kNlNL, "=LET(s;1;s+1)", "=LET(s,1,s+1)"},
  };
  for (const auto& row : kAccepted) {
    SCOPED_TRACE(row.local);
    EXPECT_EQ(Canonicalize(row.locale, row.local, At(14, 3)), row.stored);
  }
}

TEST(FormulaLocalizeR1C1, CellShapedNamesPrintBare) {
  const struct {
    ExcelLocale locale;
    const char* local;
  } kRows[] = {
      {ExcelLocale::kEnUS, "=LET(B1,5,B1+1)"},
      {ExcelLocale::kDeDE, "=LET(B1;5;B1+1)"},
      {ExcelLocale::kEsES, "=LET(B1;5;B1+1)"},
      {ExcelLocale::kNlNL, "=LET(B1;5;B1+1)"},
  };
  const std::string_view stored = "=LET(_xlpm.B1,5,_xlpm.B1+1)";
  for (const auto& row : kRows) {
    SCOPED_TRACE(row.local);
    EXPECT_EQ(Localize(row.locale, stored, NoBase()), row.local);
    EXPECT_EQ(Canonicalize(row.locale, row.local, NoBase()), stored);
  }
  EXPECT_EQ(Localize(ExcelLocale::kDeDE, "=_xlnm.B2+1", NoBase()), "=B2+1");
  EXPECT_EQ(Canonicalize(ExcelLocale::kDeDE, "=B2+1", NoBase()), "=_xlnm.B2+1");
}

TEST(FormulaLocalizeR1C1, DecimalSeparator) {
  const struct {
    ExcelLocale locale;
    const char* local;
  } kRows[] = {
      {ExcelLocale::kDeDE, "=Z(-27)S(-3)+1,5"},
      {ExcelLocale::kEsES, "=F[-27]C[-3]+1,5"},
      {ExcelLocale::kNlNL, "=R[-27]K[-3]+1,5"},
  };
  for (const auto& row : kRows) {
    SCOPED_TRACE(row.local);
    EXPECT_EQ(Localize(row.locale, "=A1+1.5", At(27, 3)), row.local);
    EXPECT_EQ(Canonicalize(row.locale, row.local, At(27, 3)), "=A1+1.5");
  }
}

TEST(FormulaLocalizeR1C1, GermanTrimReferencesRoundTrip) {
  const struct {
    std::uint32_t row;
    const char* local;
    const char* stored;
  } kRows[] = {
      {28, "=SUMME(Z(-28)S(-3):.Z(-26)S(-2))", "=SUM(A1:.B3)"},
      {29, "=SUMME(Z(-29)S(-3).:Z(-27)S(-2))", "=SUM(A1.:B3)"},
      {30, "=SUMME(Z(-30)S(-3).:.Z(-28)S(-2))", "=SUM(A1.:.B3)"},
  };
  for (const auto& row : kRows) {
    SCOPED_TRACE(row.local);
    EXPECT_EQ(Canonicalize(ExcelLocale::kDeDE, row.local, At(row.row, 3)), row.stored);
    EXPECT_EQ(Localize(ExcelLocale::kDeDE, row.stored, At(row.row, 3)), row.local);
  }
}

// ---------------------------------------------------------------------------
// Errors and positions
// ---------------------------------------------------------------------------

TEST(FormulaLocalizeR1C1, RelativeWithoutBaseFails) {
  const ExcelProfile en = Mac(ExcelLocale::kEnUS);
  const ExcelProfile de = Mac(ExcelLocale::kDeDE);
  ExpectFailure(localize_formula_text_r1c1("=SUM(A1,$B$2)", en, NoBase()), parser::ParseErrorCode::InvalidReference, 5,
                7);
  ExpectFailure(localize_formula_text_r1c1("=Sheet2!A1+1", en, NoBase()), parser::ParseErrorCode::InvalidReference, 1,
                10);
  ExpectFailure(localize_formula_text_r1c1("=$A$1+A:A", en, NoBase()), parser::ParseErrorCode::InvalidReference, 6, 9);
  ExpectFailure(canonicalize_formula_text_r1c1("=SUM(R[-1]C,R2C2)", en, NoBase()),
                parser::ParseErrorCode::InvalidReference, 5, 11);
  ExpectFailure(canonicalize_formula_text_r1c1("=1+Z(-1)S", de, NoBase()), parser::ParseErrorCode::InvalidReference, 3,
                9);
  // Absolute references and formulas without references need no base.
  EXPECT_EQ(Localize(ExcelLocale::kDeDE, "=$A$1+SUM($1:$2)", NoBase()), "=Z1S1+SUMME(Z1:Z2)");
  EXPECT_EQ(Canonicalize(ExcelLocale::kDeDE, "=Z1S1+SUMME(Z1:Z2)", NoBase()), "=$A$1+SUM($1:$2)");
  EXPECT_EQ(Localize(ExcelLocale::kDeDE, "=1.5+2", NoBase()), "=1,5+2");
  EXPECT_EQ(Canonicalize(ExcelLocale::kDeDE, "=1,5+2", NoBase()), "=1.5+2");
}

TEST(FormulaLocalizeR1C1, MalformedEndpointFails) {
  ExpectFailure(canonicalize_formula_text_r1c1("=SUMME(Z(x)S;1)", Mac(ExcelLocale::kDeDE), At(4, 4)),
                parser::ParseErrorCode::InvalidReference, 7, 11);
  ExpectFailure(canonicalize_formula_text_r1c1("=F[", Mac(ExcelLocale::kEsES), At(4, 4)),
                parser::ParseErrorCode::InvalidReference, 1, 3);
}

TEST(FormulaLocalizeR1C1, UnparsedFormulaFailsAtItsPlace) {
  const Expected<MappedFormulaText, parser::ParseFailure> out =
      localize_formula_text_r1c1("=SUM(A1,", Mac(ExcelLocale::kEnUS), At(4, 2));
  ASSERT_FALSE(out);
  EXPECT_GE(out.error().range.start, 1U);
  EXPECT_LE(out.error().range.end, 8U);
}

// In A1 the notation-taking localize fails as the entry parse does, its range
// moved past the `=`; the plain A1 rewrite still renders the same text.
TEST(FormulaLocalizeR1C1, A1NotationReportsTheParseFailure) {
  const ExcelProfile de = Mac(ExcelLocale::kDeDE);
  const std::string_view formula = "=SUM(1,";
  Arena arena;
  parser::ParseFailure expected;
  ASSERT_EQ(parse_formula_entry_with_diagnostics(strip_formula_prefix(formula), arena, &expected), nullptr);
  const Expected<MappedFormulaText, parser::ParseFailure> out = localize_formula_text_r1c1(formula, de, {});
  ASSERT_FALSE(out) << out.value().text;
  EXPECT_EQ(out.error().code, expected.code);
  EXPECT_EQ(out.error().range.start, expected.range.start + 1U);
  EXPECT_EQ(out.error().range.end, expected.range.end + 1U);
  EXPECT_EQ(out.error().range.line, expected.range.line);
  EXPECT_EQ(out.error().range.column, expected.range.line == 1U ? expected.range.column + 1U : expected.range.column);

  EXPECT_EQ(localize_formula_text(formula, de), "=SUMME(1;");
  EXPECT_EQ(localize_formula_text_mapped(formula, de).text, "=SUMME(1;");
}

TEST(FormulaLocalizeR1C1, SpansMapToTheCallersText) {
  const ExcelProfile de = Mac(ExcelLocale::kDeDE);
  const std::string_view local = "=SUMME(Z(-2)S(-2);\"x\")";
  const Expected<MappedFormulaText, parser::ParseFailure> a1 = canonicalize_formula_text_r1c1(local, de, At(4, 4));
  ASSERT_TRUE(a1);
  ASSERT_EQ(a1.value().text, "=SUM(C3,\"x\")");
  const parser::TextRange ref = a1.value().source_range(local, Units(5, 7));
  EXPECT_EQ(ref.start, 7U);
  EXPECT_EQ(ref.end, 17U);
  EXPECT_EQ(ref.column, 8U);
  const parser::TextRange text = a1.value().source_range(local, Units(8, 11));
  EXPECT_EQ(text.start, 18U);
  EXPECT_EQ(text.end, 21U);

  const std::string_view canonical = "=SUM(C3,\"x\")";
  const Expected<MappedFormulaText, parser::ParseFailure> back = localize_formula_text_r1c1(canonical, de, At(4, 4));
  ASSERT_TRUE(back);
  ASSERT_EQ(back.value().text, local);
  const parser::TextRange cell = back.value().source_range(canonical, Units(7, 17));
  EXPECT_EQ(cell.start, 5U);
  EXPECT_EQ(cell.end, 7U);
}

// ---------------------------------------------------------------------------
// Round trip and agreement with the AST printer
// ---------------------------------------------------------------------------

TEST(FormulaLocalizeR1C1, RoundTripsEveryProfile) {
  for (const ExcelLocale locale : kAllLocales) {
    for (const ProbeRow& row : kProbeRows) {
      const std::string a1 = std::string("=") + row.a1;
      SCOPED_TRACE(testing::Message() << static_cast<int>(locale) << " " << a1);
      const std::string local = Localize(locale, a1, At(row.row, row.col));
      EXPECT_EQ(Canonicalize(locale, local, At(row.row, row.col)), a1) << local;
    }
  }
}

// The A1 localization of the invariant AST printer's text.
std::string ViaAstPrinter(std::string_view a1, std::uint32_t row, std::uint32_t col) {
  Arena arena;
  const parser::AstNode* root = parse_formula_entry(a1, arena);
  if (root == nullptr) {
    return "<unparsed>";
  }
  return parser::format_formula_r1c1(*root, row, col);
}

// en-US localizes no token, so its R1C1 output is the AST printer's text.
TEST(FormulaLocalizeR1C1, AgreesWithTheAstPrinter) {
  for (const ProbeRow& row : kProbeRows) {
    SCOPED_TRACE(row.a1);
    EXPECT_EQ(Localize(ExcelLocale::kEnUS, row.a1, At(row.row, row.col)), ViaAstPrinter(row.a1, row.row, row.col));
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
