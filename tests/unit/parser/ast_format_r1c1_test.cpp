#include "parser/ast_format_r1c1.h"

#include <cstdint>
#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "parser/reference.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

// Parses `src` (no leading `=`) and formats it as R1C1 at the 0-based host cell.
std::string R1C1(std::string_view src, std::uint32_t row, std::uint32_t col) {
  Arena arena;
  Parser p(src, arena);
  AstNode* root = p.parse();
  EXPECT_TRUE(p.errors().empty()) << "parse errors for: " << src;
  if (root == nullptr) {
    return "<null>";
  }
  return format_formula_r1c1(*root, row, col);
}

// Rows measured against Mac Excel `FormulaR1C1`: `a1` is the normalized (read
// back) A1 spelling, `r1c1` Excel's answer, both without the leading `=`.
struct ProbeRow {
  const char* a1;
  std::uint32_t row;
  std::uint32_t col;
  const char* r1c1;
};

// The probe's `[o7_ext.xlsx]` rows are entered (and expected) with the book
// spelled `[1]`, which reads the same way; the external whole-column row is
// also built by hand below.
constexpr ProbeRow kProbeRows[] = {
    {"A1", 4, 2, "R[-4]C[-2]"},
    {"$A$1", 4, 2, "R1C1"},
    {"$A1", 4, 2, "R[-4]C1"},
    {"A$1", 4, 2, "R1C[-2]"},
    {"C5", 4, 2, "RC"},
    {"D6", 4, 2, "R[1]C[1]"},
    {"B4", 4, 2, "R[-1]C[-1]"},
    {"A1:B2", 4, 2, "R[-4]C[-2]:R[-3]C[-1]"},
    {"$A$1:B2", 4, 2, "R1C1:R[-3]C[-1]"},
    {"A1:$B$2", 4, 2, "R[-4]C[-2]:R2C2"},
    {"SUM(A:A)", 4, 2, "SUM(C[-2])"},
    {"SUM($A:$A)", 4, 2, "SUM(C1)"},
    {"SUM(A:C)", 4, 2, "SUM(C[-2]:C)"},
    {"SUM(D:E)", 4, 2, "SUM(C[1]:C[2])"},
    {"SUM($A:E)", 4, 2, "SUM(C1:C[2])"},
    {"SUM(1:1)", 4, 2, "SUM(R[-4])"},
    {"SUM($1:$1)", 4, 2, "SUM(R1)"},
    {"SUM(4:6)", 4, 2, "SUM(R[-1]:R[1])"},
    {"SUM($4:6)", 4, 2, "SUM(R4:R[1])"},
    {"Data!A1", 4, 2, "Data!R[-4]C[-2]"},
    {"Data!$A$1", 4, 2, "Data!R1C1"},
    {"Data!A1:B2", 4, 2, "Data!R[-4]C[-2]:R[-3]C[-1]"},
    {"'My Sheet'!A1", 4, 2, "'My Sheet'!R[-4]C[-2]"},
    {"'It''s'!A1", 4, 2, "'It''s'!R[-4]C[-2]"},
    {"'My Sheet'!$B$2:C3", 4, 2, "'My Sheet'!R2C2:R[-2]C"},
    {"Data!A:A", 4, 2, "Data!C[-2]"},
    {"Data!1:1", 4, 2, "Data!R[-4]"},
    {"Data:'My Sheet'!A1", 4, 2, "Data:'My Sheet'!R[-4]C[-2]"},
    {"SUM(Data:'My Sheet'!A1:B2)", 4, 2, "SUM(Data:'My Sheet'!R[-4]C[-2]:R[-3]C[-1])"},
    {"Tbl[@Col1]", 4, 2, "Tbl[@Col1]"},
    {"SUM(Tbl[@[Col1]:[Col2]])", 4, 2, "SUM(Tbl[@[Col1]:[Col2]])"},
    {"Tbl[@Col1]", 4, 2, "Tbl[@Col1]"},
    {"SUM(Tbl[#Data])", 4, 2, "SUM(Tbl[#Data])"},
    {"Tbl[#Headers]", 4, 2, "Tbl[#Headers]"},
    {"Tbl[[#Totals],[Col2]]", 4, 2, "Tbl[[#Totals],[Col2]]"},
    {"MyName", 4, 2, "MyName"},
    {"MyName+A1", 4, 2, "MyName+R[-4]C[-2]"},
    {"SUM(MyRange)", 4, 2, "SUM(MyRange)"},
    {"Data!LocalName", 4, 2, "Data!LocalName"},
    {"MyRel", 4, 2, "MyRel"},
    {"\"A1\"", 4, 2, "\"A1\""},
    {"\"R1C1 and A1:B2\"&A1", 4, 2, "\"R1C1 and A1:B2\"&R[-4]C[-2]"},
    {"IF(A1=\"B2\",1,0)", 4, 2, "IF(R[-4]C[-2]=\"B2\",1,0)"},
    {"\"Sheet1!A1\"", 4, 2, "\"Sheet1!A1\""},
    {"\"He said \"\"A1\"\"\"", 4, 2, "\"He said \"\"A1\"\"\""},
    {"A1+B2*C3", 4, 2, "R[-4]C[-2]+R[-3]C[-1]*R[-2]C"},
    {"SUM(A1,B2,$C$3)", 4, 2, "SUM(R[-4]C[-2],R[-3]C[-1],R3C3)"},
    {"INDEX(A1:C3,2,2)", 4, 2, "INDEX(R[-4]C[-2]:R[-2]C,2,2)"},
    {"A1:B2 B2:C3", 4, 2, "R[-4]C[-2]:R[-3]C[-1] R[-3]C[-1]:R[-2]C"},
    {"(A1,B2)", 4, 2, "(R[-4]C[-2],R[-3]C[-1])"},
    {"A1:INDEX(B:B,3)", 4, 2, "R[-4]C[-2]:INDEX(C[-1],3)"},
    {"$A$1", 4, 2, "R1C1"},
    {"C5", 4, 2, "RC"},
    {"XFD1048576", 4, 2, "R[1048571]C[16381]"},
    {"A1048576", 4, 2, "R[1048571]C[-2]"},
    {"XFD1", 4, 2, "R[-4]C[16381]"},
    {"A1#", 4, 2, "R[-4]C[-2]#"},
    {"A1:A3", 4, 2, "R[-4]C[-2]:R[-2]C[-2]"},
    {"LET(x,A1,x+B2)", 4, 2, "LET(x,R[-4]C[-2],x+R[-3]C[-1])"},
    {"LAMBDA(a,a+A1)(2)", 4, 2, "LAMBDA(a,a+R[-4]C[-2])(2)"},
    {"B2", 0, 0, "R[1]C[1]"},
    {"Z100", 0, 0, "R[99]C[25]"},
    {"A1", 1048575, 16383, "R[-1048575]C[-16383]"},
    {"XFD1048575", 1048575, 16383, "R[-1]C"},
    {"A1", 1, 1, "R[-1]C[-1]"},
    {"SUM(A:A)", 1, 1, "SUM(C[-1])"},
    {"[1]Sheet1!A1", 4, 2, "[1]Sheet1!R[-4]C[-2]"},
    {"[1]Sheet1!$A$1:B2", 4, 2, "[1]Sheet1!R1C1:R[-3]C[-1]"},
    {"'[1]It''s'!A1", 4, 2, "'[1]It''s'!R[-4]C[-2]"},
};

TEST(AstFormatR1C1, ProbeTable) {
  for (const ProbeRow& row : kProbeRows) {
    EXPECT_EQ(R1C1(row.a1, row.row, row.col), row.r1c1) << "A1: " << row.a1;
  }
}

TEST(AstFormatR1C1, ExternalWholeColumn) {
  Arena arena;
  Reference col;
  col.col = 0;
  col.is_full_col = true;
  const AstNode* node = make_external_ref(arena, {}, "1", "Sheet1", {}, col, col, false);
  EXPECT_EQ(format_formula_r1c1(*node, 4, 2), "[1]Sheet1!C[-2]");
}

TEST(AstFormatR1C1, HostCellIsRC) {
  EXPECT_EQ(R1C1("C5", 4, 2), "RC");
  EXPECT_EQ(R1C1("B2*(A1+C3)", 1, 1), "RC*(R[-1]C[-1]+R[1]C[1])");
}

TEST(AstFormatR1C1, AbsoluteAndMixed) {
  EXPECT_EQ(R1C1("$A$1", 4, 2), "R1C1");
  EXPECT_EQ(R1C1("$A1", 4, 2), "R[-4]C1");
  EXPECT_EQ(R1C1("A$1", 4, 2), "R1C[-2]");
}

TEST(AstFormatR1C1, ReferenceLikeTextInStringsIsUntouched) {
  EXPECT_EQ(R1C1("\"A1:B2\"&A1", 4, 2), "\"A1:B2\"&R[-4]C[-2]");
}

}  // namespace
}  // namespace parser
}  // namespace formulon
