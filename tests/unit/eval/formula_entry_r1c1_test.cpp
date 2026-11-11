//
// Unit tests for the R1C1 entry fallback: the three-way endpoint reader in
// `eval/a1_parse.h`, the triggers that send A1-parsed entry to R1C1, and
// the lexical R1C1 -> A1 rewrite.

#include "eval/formula_entry_r1c1.h"

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

#include "eval/a1_parse.h"
#include "eval/builtin_names.h"
#include "eval/formula_localize.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parse_error.h"
#include "parser/parse_status.h"
#include "table.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/formula_text.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using refs_internal::parse_r1c1_endpoint_axes;
using refs_internal::R1C1Endpoint;
using refs_internal::R1C1EndpointKind;
using refs_internal::R1C1Letters;

R1C1EndpointKind KindOf(std::string_view text) {
  return parse_r1c1_endpoint_axes(text, 0, R1C1Letters{}).kind;
}

std::string Translate(std::string_view text, std::uint32_t row = 0, std::uint32_t col = 0) {
  const std::optional<std::string> a1 = translate_r1c1_entry(text, row, col);
  return a1 ? *a1 : std::string("<refused>");
}

bool NeedsR1C1(std::string_view text, const Workbook& book) {
  Arena arena;
  const parser::AstNode* root = parse_formula_entry(strip_formula_prefix(text), arena);
  EXPECT_NE(root, nullptr) << text;
  return root != nullptr && needs_r1c1_entry(*root, text, book);
}

// ---------------------------------------------------------------------------
// parse_r1c1_endpoint_axes
// ---------------------------------------------------------------------------

TEST(FormulaEntryR1C1, EndpointInsideTheGridIsAReference) {
  for (const char* text :
       {"R1C1", "R1048576", "C16384", "R", "C", "RC", "R[1048576]", "R[-1048576]C[+16384]", "r1c1", "rc[-1]", "R001"}) {
    SCOPED_TRACE(text);
    EXPECT_EQ(KindOf(text), R1C1EndpointKind::kReference);
  }
}

TEST(FormulaEntryR1C1, AbsoluteAxisOutsideTheGridIsAName) {
  for (const char* text : {"R0", "R1048577", "C0", "C16385", "R4294967296", "R1C16385", "R99999999999999999999"}) {
    SCOPED_TRACE(text);
    EXPECT_EQ(KindOf(text), R1C1EndpointKind::kName);
  }
}

TEST(FormulaEntryR1C1, BadBracketIsMalformed) {
  for (const char* text :
       {"R[1048577]", "R[4294967296]", "C[-1048577]", "R[", "R[]", "R[x]", "R[1", "R[1]C0", "X1", ""}) {
    SCOPED_TRACE(text);
    EXPECT_EQ(KindOf(text), R1C1EndpointKind::kMalformed);
  }
}

TEST(FormulaEntryR1C1, EndpointReportsAxesAndWhereItStops) {
  const R1C1Endpoint ep = parse_r1c1_endpoint_axes("=R5C[-2]+1", 1, R1C1Letters{});
  EXPECT_EQ(ep.kind, R1C1EndpointKind::kReference);
  EXPECT_TRUE(ep.bracketed);
  EXPECT_EQ(ep.end, 8U);
  EXPECT_TRUE(ep.row.present);
  EXPECT_FALSE(ep.row.relative);
  EXPECT_EQ(ep.row.value, 5);
  EXPECT_TRUE(ep.col.present);
  EXPECT_TRUE(ep.col.relative);
  EXPECT_EQ(ep.col.value, -2);

  const R1C1Endpoint row_only = parse_r1c1_endpoint_axes("R1C1C", 0, R1C1Letters{});
  EXPECT_EQ(row_only.end, 4U);
  EXPECT_FALSE(row_only.bracketed);

  const R1C1Endpoint de = parse_r1c1_endpoint_axes("Z(1)S", 0, R1C1Letters{'Z', 'S', '(', ')'});
  EXPECT_EQ(de.kind, R1C1EndpointKind::kReference);
  EXPECT_EQ(de.end, 5U);
}

// ---------------------------------------------------------------------------
// translate_r1c1_entry: references
// ---------------------------------------------------------------------------

TEST(FormulaEntryR1C1, CellEndpointsBecomeA1) {
  EXPECT_EQ(Translate("=RC[-1]", 2, 4), "=D3");
  EXPECT_EQ(Translate("RC[-1]", 2, 4), "D3");
  EXPECT_EQ(Translate("=R1C1"), "=$A$1");
  EXPECT_EQ(Translate("=R[0]C[-1]", 4, 4), "=D5");
  EXPECT_EQ(Translate("=R2C[1]", 4, 4), "=F$2");
  EXPECT_EQ(Translate("=r1c1+rc[-1]", 16, 4), "=$A$1+D17");
  EXPECT_EQ(Translate("=Sheet1!R1C1"), "=Sheet1!$A$1");
  EXPECT_EQ(Translate("='R1C1'!R1C1"), "='R1C1'!$A$1");
  EXPECT_EQ(Translate("=R1C1#"), "=$A$1#");
}

TEST(FormulaEntryR1C1, RelativeAxisWrapsAroundTheGrid) {
  EXPECT_EQ(Translate("=RC[-10]", 12, 4), "=XEY13");
  EXPECT_EQ(Translate("=R[1]C", 1048575, 0), "=A1");
  EXPECT_EQ(Translate("=R[-1048576]C", 0, 0), "=A1");
  EXPECT_EQ(Translate("=RC[1]", 0, 16383), "=A1");
}

TEST(FormulaEntryR1C1, LoneRowOrColumnSpansItself) {
  EXPECT_EQ(Translate("=COUNT(R1)"), "=COUNT($1:$1)");
  EXPECT_EQ(Translate("=COLUMNS(R[1])", 11, 4), "=COLUMNS(13:13)");
  EXPECT_EQ(Translate("=SUM(C1)"), "=SUM($A:$A)");
  EXPECT_EQ(Translate("=ROWS(C[-1])", 10, 4), "=ROWS(D:D)");
  EXPECT_EQ(Translate("=R", 4, 4), "=5:5");
  EXPECT_EQ(Translate("=C", 4, 4), "=E:E");
}

TEST(FormulaEntryR1C1, RangesJoinWholeAxesAndKeepCellEndpoints) {
  EXPECT_EQ(Translate("=SUM(R1:R2)"), "=SUM($1:$2)");
  EXPECT_EQ(Translate("=SUM(C1:C2)"), "=SUM($A:$B)");
  EXPECT_EQ(Translate("=SUM(R[1]:R[2])", 9, 4), "=SUM(11:12)");
  EXPECT_EQ(Translate("=SUM(R1 : R2)"), "=SUM($1 : $2)");
  EXPECT_EQ(Translate("=R1C1:R2C1"), "=$A$1:$A$2");
  EXPECT_EQ(Translate("=SUM(R1C1 :R2C1)"), "=SUM($A$1 :$A$2)");
  EXPECT_EQ(Translate("=Sheet1!R1C1:R2C1"), "=Sheet1!$A$1:$A$2");
}

TEST(FormulaEntryR1C1, AbsoluteAxisOutsideTheGridStaysAName) {
  EXPECT_EQ(Translate("=LET(B1,5,B1+R1048577)"), "=LET(_xlpm.B1,5,_xlpm.B1+R1048577)");
  EXPECT_EQ(Translate("=LET(B1,5,B1+R0)"), "=LET(_xlpm.B1,5,_xlpm.B1+R0)");
  EXPECT_EQ(Translate("=LET(B1,5,B1+R4294967296)"), "=LET(_xlpm.B1,5,_xlpm.B1+R4294967296)");
  EXPECT_EQ(Translate("=LET(B1,5,B1+C16385)"), "=LET(_xlpm.B1,5,_xlpm.B1+_xlnm.C16385)");
}

TEST(FormulaEntryR1C1, RefusesWhatR1C1EntryCannotHold) {
  for (const char* text :
       {"=LET(B1,5,B1+$B1)", "=LET(B1,5,B1+$A$1)", "=LET(B1,5,SUM(1:2))", "=SUM(1 :2)", "=R[-5000000]C[-1]",
        "=R[1048577]", "=RC[x]", "=LET(RC,1,RC)+R1C1", "=LET(B1,5,R1,6,B1+R1)", "=LET(R1C,2,R1C)+R1C1", "=SUM("}) {
    SCOPED_TRACE(text);
    EXPECT_EQ(Translate(text, 4, 4), "<refused>");
  }
}

// ---------------------------------------------------------------------------
// translate_r1c1_entry: names and verbatim runs
// ---------------------------------------------------------------------------

TEST(FormulaEntryR1C1, CellShapedNamesTakeTheirStoredPrefix) {
  EXPECT_EQ(Translate("=LET(B1,5,B1+R1C1)"), "=LET(_xlpm.B1,5,_xlpm.B1+$A$1)");
  EXPECT_EQ(Translate("=B2+LET(B1,5,B1)"), "=_xlnm.B2+LET(_xlpm.B1,5,_xlpm.B1)");
  EXPECT_EQ(Translate("=LET(B1,5,B1+A1:B2)"), "=LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1:_xlnm.B2)");
  EXPECT_EQ(Translate("=LAMBDA(B1,B1+R1C1)(4)"), "=LAMBDA(_xlpm.B1,_xlpm.B1+$A$1)(4)");
  EXPECT_EQ(Translate("=LET(x,1,x)+LET(B1,5,B1)"), "=LET(x,1,x)+LET(_xlpm.B1,5,_xlpm.B1)");
  // A LET value sees only the names bound before it.
  EXPECT_EQ(Translate("=LET(x,B1,B1,2,B1+x)"), "=LET(x,_xlnm.B1,_xlpm.B1,2,_xlpm.B1+x)");
}

TEST(FormulaEntryR1C1, SheetQualifiedCellShapeIsAnUndefinedName) {
  EXPECT_EQ(Translate("=LET(B1,5,Sheet1!B1)"), "=LET(_xlpm.B1,5,Sheet1!_xlnm.B1)");
  EXPECT_EQ(Translate("=LET(B1,5,SUM(Sheet1!A1:B2))"), "=LET(_xlpm.B1,5,SUM(Sheet1!_xlnm.A1:_xlnm.B2))");
}

TEST(FormulaEntryR1C1, CellShapedCalleeIsANameUnlessBuiltIn) {
  EXPECT_EQ(Translate("=LET(B1,100,LOG10(B1))"), "=LET(_xlpm.B1,100,LOG10(_xlpm.B1))");
  EXPECT_EQ(Translate("=LET(B1,LAMBDA(x,x+1),B1(2))"), "=LET(_xlpm.B1,LAMBDA(x,x+1),_xlpm.B1(2))");
  EXPECT_EQ(Translate("=LAMBDA(B1,B1(1))(LAMBDA(x,x*2))"), "=LAMBDA(_xlpm.B1,_xlpm.B1(1))(LAMBDA(x,x*2))");
  EXPECT_EQ(Translate("=B2(1)+R1C1"), "=_xlnm.B2(1)+$A$1");
}

TEST(FormulaEntryR1C1, ColumnShapedNameRightOfANameColonIsParenthesised) {
  EXPECT_EQ(Translate("=LET(B1,5,B1+A:A)"), "=LET(_xlpm.B1,5,_xlpm.B1+A:(A))");
  EXPECT_EQ(Translate("=LET(B1,5,SUM(A:B))"), "=LET(_xlpm.B1,5,SUM(A:(B)))");
  EXPECT_EQ(Translate("=LET(B1,5,B1+A1:A)"), "=LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1:(A))");
  EXPECT_EQ(Translate("=LET(B1,5,B1+A)"), "=LET(_xlpm.B1,5,_xlpm.B1+A)");
  EXPECT_EQ(Translate("=LET(B1,5,R1C1:B1)"), "=LET(_xlpm.B1,5,$A$1:_xlpm.B1)");
}

TEST(FormulaEntryR1C1, StringsQuotedNamesAndBracketsAreVerbatim) {
  EXPECT_EQ(Translate("=LET(B1,5,INDIRECT(\"R1C1\",FALSE)+B1)"), "=LET(_xlpm.B1,5,INDIRECT(\"R1C1\",FALSE)+_xlpm.B1)");
  EXPECT_EQ(Translate("=LEN(\"$B1 \"\"R1\"\"\")+R1C1"), "=LEN(\"$B1 \"\"R1\"\"\")+$A$1");
  EXPECT_EQ(Translate("=Tbl[RC]+R1C1"), "=Tbl[RC]+$A$1");
  EXPECT_EQ(Translate("=Tbl[[#This Row],[B1]]+R1C1"), "=Tbl[[#This Row],[B1]]+$A$1");
  EXPECT_EQ(Translate("=[Book.xlsx]Sheet1!R1C1"), "=[Book.xlsx]Sheet1!$A$1");
  EXPECT_EQ(Translate("=IF(ISNA(R1C1),#N/A,R1C1)"), "=IF(ISNA($A$1),#N/A,$A$1)");
  EXPECT_EQ(Translate("=LET(B1,5,B1+XFE1+名前)"), "=LET(_xlpm.B1,5,_xlpm.B1+XFE1+名前)");
  EXPECT_EQ(Translate("=LET(B1,5,B1+_xlnm.B2)"), "=LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2)");
}

TEST(FormulaEntryR1C1, WhitespaceIsKept) {
  EXPECT_EQ(Translate("=LET(B1,5,B1 + R1C1)"), "=LET(_xlpm.B1,5,_xlpm.B1 + $A$1)");
  EXPECT_EQ(Translate("= LET( B1 , 5 , B1+B2 )"), "= LET( _xlpm.B1 , 5 , _xlpm.B1+_xlnm.B2 )");
  EXPECT_EQ(Translate("=SUM( RC[-1] , 1 )", 12, 4), "=SUM( D13 , 1 )");
  EXPECT_EQ(Translate("=RC[-1] + 1", 11, 4), "=D12 + 1");
}

// ---------------------------------------------------------------------------
// rewrite_r1c1_text outside entry
// ---------------------------------------------------------------------------

R1C1RewriteOptions Rewriting(const R1C1Letters& letters, char array_row_separator, std::uint32_t row,
                             std::uint32_t col) {
  R1C1RewriteOptions options;
  options.letters = letters;
  options.array_row_separator = array_row_separator;
  options.has_base = true;
  options.base_row = row;
  options.base_col = col;
  return options;
}

std::string Rewrite(std::string_view text, const R1C1RewriteOptions& options) {
  const Expected<MappedFormulaText, parser::ParseFailure> a1 = rewrite_r1c1_text(text, options);
  return a1 ? a1.value().text : std::string("<failed>");
}

TEST(FormulaEntryR1C1, RewriteReadsTheGivenLetters) {
  const R1C1RewriteOptions de = Rewriting(R1C1Letters{'Z', 'S', '(', ')'}, ';', 4, 4);
  EXPECT_EQ(Rewrite("=SUMME(Z(-2)S(-2);Z2S2)", de), "=SUMME(C3;$B$2)");
  EXPECT_EQ(Rewrite("=SUMME(S1:S2)+ZEILE(Z)", de), "=SUMME($A:$B)+ZEILE(5:5)");
  // The invariant letters are names under other letters; a cell-shaped name takes `_xlpm.`.
  EXPECT_EQ(Rewrite("=R1C1+B1", de), "=R1C1+_xlpm.B1");
  EXPECT_EQ(Rewrite("=Tbl[Col]+\"Z1S1\"+'Z1'!Z1S1", de), "=Tbl[Col]+\"Z1S1\"+'Z1'!$A$1");
  // Qualifiers take the stored quoting rule; a partly quoted 3-D span stays as written.
  EXPECT_EQ(Rewrite("=S2!Z1S1+'Data'!Z1S1+Data:S2!Z1S1+Data:'My Sheet'!Z1S1", de),
            "='S2'!$A$1+Data!$A$1+'Data:S2'!$A$1+Data:'My Sheet'!$A$1");
  EXPECT_EQ(Rewrite("=[1]S2!Z1S1+'[1]S2'!Z1S1", de), "=[1]S2!$A$1+'[1]S2'!$A$1");
}

TEST(FormulaEntryR1C1, RewriteOutsideEntryKeepsArrayRows) {
  const R1C1RewriteOptions ru = Rewriting(R1C1Letters{}, ':', 4, 4);
  EXPECT_EQ(Rewrite("=SUM(R1C1;{1;2:3;4})", ru), "=SUM($A$1;{1;2:3;4})");
  // Entry refuses a number before `:` even inside braces.
  EXPECT_EQ(Translate("={1,2:3,4}"), "<refused>");
}

TEST(FormulaEntryR1C1, RewriteJoinsTrimOperatorsOutsideEntry) {
  const R1C1RewriteOptions en = Rewriting(R1C1Letters{}, ';', 4, 2);
  EXPECT_EQ(Rewrite("=SUM(R[-4]C[-2].:R[-3]C[-1])", en), "=SUM(A1.:B2)");
  EXPECT_EQ(Rewrite("=SUM(C[-2]:.C[-2])", en), "=SUM(A:.A)");
  EXPECT_EQ(Rewrite("=SUM(R1.:.R3)", en), "=SUM($1.:.$3)");
}

TEST(FormulaEntryR1C1, RewriteFailuresCarryTheEndpoint) {
  R1C1RewriteOptions no_base;
  const Expected<MappedFormulaText, parser::ParseFailure> relative = rewrite_r1c1_text("=1+RC[-1]", no_base);
  ASSERT_FALSE(relative);
  EXPECT_EQ(relative.error().code, parser::ParseErrorCode::InvalidReference);
  EXPECT_EQ(relative.error().range.start, 3U);
  EXPECT_EQ(relative.error().range.end, 9U);
  const Expected<MappedFormulaText, parser::ParseFailure> malformed = rewrite_r1c1_text("=R[x]+1", no_base);
  ASSERT_FALSE(malformed);
  EXPECT_EQ(malformed.error().code, parser::ParseErrorCode::InvalidReference);
  EXPECT_EQ(malformed.error().range.start, 1U);
  EXPECT_EQ(malformed.error().range.end, 5U);
}

// ---------------------------------------------------------------------------
// needs_r1c1_entry and has_entered_cell_shaped_param
// ---------------------------------------------------------------------------

TEST(FormulaEntryR1C1, UnboundR1C1ShapedNameTriggersTheFallback) {
  const Workbook book = Workbook::create();
  for (const char* text : {"=R1C1+1", "=LET(x,5,x+R1C1)", "=Sheet1!R1C1", "=RC", "=R", "=r1c1", "=LET(R1C,2,R1C)+R1C1",
                           "=LAMBDA(x,x+RC)(1)", "=LET(RC,1,Sheet1!RC)"}) {
    SCOPED_TRACE(text);
    EXPECT_TRUE(NeedsR1C1(text, book));
  }
}

TEST(FormulaEntryR1C1, BracketedEndpointTriggersTheFallback) {
  const Workbook book = Workbook::create();
  for (const char* text : {"=RC[-1]", "=rc[-1]", "=SUM(R[1]:R[2])", "=ROWS(C[-1])", "=COLUMNS(R[1])", "=R1C[2]+1"}) {
    SCOPED_TRACE(text);
    EXPECT_TRUE(NeedsR1C1(text, book));
  }
}

TEST(FormulaEntryR1C1, A1EntryWithoutR1C1ShapesKeepsItsParse) {
  const Workbook book = Workbook::create();
  for (const char* text : {"=LET(R1C1,3,R1C1)", "=LAMBDA(R1C,R1C)(3)", "=LET(R1C,2,R1C)", "=R1C1C", "=R0+R1048577",
                           "=SUM(R1:R2)", "=\"R1C1\"&\"RC[-1]\"", "=INDIRECT(\"RC[-1]\",FALSE)", "=Tbl[Qty]",
                           "=Rates[Qty]", "=LET(x,5,x+B2)", "=R1C1!A1", "=_xlnm.B2"}) {
    SCOPED_TRACE(text);
    EXPECT_FALSE(NeedsR1C1(text, book));
  }
}

TEST(FormulaEntryR1C1, TableNameBeforeTheBracketKeepsTheStructuredReference) {
  Workbook book = Workbook::create();
  TableMetadata table;
  table.id = 1;
  table.name = "RC";
  table.display_name = "RC";
  table.ref = "A1:A3";
  table.columns.emplace_back(1U, "Qty", "", "", "");
  book.mutable_tables().push_back(table);
  EXPECT_FALSE(NeedsR1C1("=SUM(RC[Qty])", book));
  EXPECT_FALSE(NeedsR1C1("=SUM(rc[Qty])", book));
}

TEST(FormulaEntryR1C1, TypedParamPrefixOnACellShapeIsDetected) {
  EXPECT_TRUE(has_entered_cell_shaped_param("=LET(_xlpm.B1,5,_xlpm.B1)"));
  EXPECT_TRUE(has_entered_cell_shaped_param("=1+_XLPM.b1"));
  EXPECT_FALSE(has_entered_cell_shaped_param("=LET(_xlpm.x,5,_xlpm.x)"));
  EXPECT_FALSE(has_entered_cell_shaped_param("=\"_xlpm.B1\""));
  EXPECT_FALSE(has_entered_cell_shaped_param("='_xlpm.B1'!A1"));
  EXPECT_FALSE(has_entered_cell_shaped_param("=_xlnm.B2"));
  EXPECT_FALSE(has_entered_cell_shaped_param("=a_xlpm.B1"));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
