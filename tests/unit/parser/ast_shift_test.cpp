// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the relative-reference shifter. Coverage targets the
// per-NodeKind walk: each kind that may carry a Reference is exercised
// directly, with absolute / relative axis combinations, whole-column
// and whole-row endpoints, and out-of-bounds shifts that collapse to
// `#REF!`. The S-expression dumper provides the golden contract so
// these tests document the visible shift behaviour.

#include "parser/ast_shift.h"

#include <cstdint>
#include <string>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_dump.h"
#include "parser/parser.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace parser {
namespace {

const AstNode* ParseSource(Arena& arena, const std::string& source) {
  Parser parser(source, arena);
  return parser.parse();
}

std::string ShiftAndDump(Arena& arena, const std::string& source, std::int32_t row_delta, std::int32_t col_delta) {
  const AstNode* root = ParseSource(arena, source);
  if (root == nullptr) {
    return "<parse-failed>";
  }
  const AstNode* shifted = shift_relative_refs(*root, arena, row_delta, col_delta);
  if (shifted == nullptr) {
    return "<shift-failed>";
  }
  return dump_sexpr(*shifted);
}

// ---------------------------------------------------------------------------
// Single Ref
// ---------------------------------------------------------------------------

TEST(AstShift, RelativeRefShiftsRowAndColumn) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "A1", 2, 1), "(ref B3)");
}

TEST(AstShift, RelativeRefNegativeShift) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "C5", -2, -1), "(ref B3)");
}

TEST(AstShift, AbsoluteBothPreserved) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "$A$1", 5, 5), "(ref $A$1)");
}

TEST(AstShift, MixedColAbsoluteRowRelative) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "$A1", 3, 4), "(ref $A4)");
}

TEST(AstShift, MixedRowAbsoluteColRelative) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "A$1", 3, 4), "(ref E$1)");
}

TEST(AstShift, ZeroDeltaIsIdentity) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "B7", 0, 0), "(ref B7)");
}

// ---------------------------------------------------------------------------
// Out-of-bounds collapses to ErrorLiteral(#REF!)
// ---------------------------------------------------------------------------

TEST(AstShift, NegativeRowOutOfBoundsBecomesRefError) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "A1", -1, 0), "(err-lit #REF!)");
}

TEST(AstShift, NegativeColOutOfBoundsBecomesRefError) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "A1", 0, -1), "(err-lit #REF!)");
}

TEST(AstShift, AbsoluteRefIsImmuneToOutOfBoundsShift) {
  Arena a;
  // $A$1 with a (-100, -100) shift is unchanged because both axes are
  // absolute.
  EXPECT_EQ(ShiftAndDump(a, "$A$1", -100, -100), "(ref $A$1)");
}

TEST(AstShift, OneAxisAbsoluteOnlyChecksRelativeAxis) {
  Arena a;
  // $A1 with row_delta=-1 lands at row 0 so it's fine; col is locked.
  EXPECT_EQ(ShiftAndDump(a, "$A2", -1, 5), "(ref $A1)");
  // But row_delta=-2 from row 2 lands at row 0 → out of bounds.
  EXPECT_EQ(ShiftAndDump(a, "$A1", -1, 5), "(err-lit #REF!)");
}

// ---------------------------------------------------------------------------
// Whole-column / whole-row references
// ---------------------------------------------------------------------------

TEST(AstShift, WholeColumnShiftsHorizontallyOnly) {
  Arena a;
  // A:A with (5, 1) shifts column to B; row delta is ignored.
  EXPECT_EQ(ShiftAndDump(a, "A:A", 5, 1), "(ref B:B)");
}

TEST(AstShift, WholeColumnAbsoluteUnaffected) {
  // The parser surface for `$A:$A` is fragile; build the Reference
  // directly so the shifter is the only thing under test.
  Arena a;
  Reference r{};
  r.is_full_col = true;
  r.col = 0;
  r.col_abs = true;
  AstNode* node = make_ref(a, r);
  ASSERT_NE(node, nullptr);
  const AstNode* shifted = shift_relative_refs(*node, a, 0, 7);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(dump_sexpr(*shifted), "(ref $A:$A)");
}

TEST(AstShift, WholeRowShiftsVerticallyOnly) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "1:1", 4, 100), "(ref 5:5)");
}

TEST(AstShift, WholeColumnNegativeOutOfBoundsBecomesRefError) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "A:A", 0, -1), "(err-lit #REF!)");
}

// ---------------------------------------------------------------------------
// RangeOp endpoints
// ---------------------------------------------------------------------------

TEST(AstShift, RangeBothEndpointsShift) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "A1:B2", 1, 1), "(range (ref B2) (ref C3))");
}

TEST(AstShift, RangeMixedAbsoluteRelative) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "$A$1:B2", 1, 1), "(range (ref $A$1) (ref C3))");
}

// ---------------------------------------------------------------------------
// BinaryOp / UnaryOp wrap children
// ---------------------------------------------------------------------------

TEST(AstShift, BinaryOpRecursesIntoChildren) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "A1>10", 1, 0), "(binary > (ref A2) (num 10))");
}

TEST(AstShift, UnaryNegationRecursesIntoOperand) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "-A1", 1, 1), "(unary - (ref B2))");
}

TEST(AstShift, ComparisonChainShiftsBothSides) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "A1>B1", 0, 0), "(binary > (ref A1) (ref B1))");
  EXPECT_EQ(ShiftAndDump(a, "A1>B1", 2, 3), "(binary > (ref D3) (ref E3))");
}

// ---------------------------------------------------------------------------
// Function calls
// ---------------------------------------------------------------------------

TEST(AstShift, CallArgsShift) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "SUM(A1:A3)", 1, 0), "(call SUM (range (ref A2) (ref A4)))");
}

TEST(AstShift, CallWithMixedLiteralAndRef) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "IF(A1>0,B1,C1)", 1, 0), "(call IF (binary > (ref A2) (num 0)) (ref B2) (ref C2))");
}

TEST(AstShift, ZeroArityCall) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "NOW()", 5, 5), "(call NOW)");
}

// ---------------------------------------------------------------------------
// Sheet-qualified references still shift
// ---------------------------------------------------------------------------

TEST(AstShift, SheetQualifiedRelativeShifts) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "Sheet2!A1", 1, 1), "(ref Sheet2!B2)");
}

TEST(AstShift, SheetQualifiedAbsoluteUnaffected) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "Sheet2!$A$1", 1, 1), "(ref Sheet2!$A$1)");
}

// ---------------------------------------------------------------------------
// Literals and other ref-free leaves are returned (logically) unchanged
// ---------------------------------------------------------------------------

TEST(AstShift, NumberLiteralUnchanged) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "42", 5, 7), "(num 42)");
}

TEST(AstShift, StringLiteralUnchanged) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "\"hello\"", 5, 7), "(text \"hello\")");
}

TEST(AstShift, BoolLiteralUnchanged) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "TRUE", 5, 7), "(bool true)");
}

TEST(AstShift, ErrorLiteralUnchanged) {
  Arena a;
  EXPECT_EQ(ShiftAndDump(a, "#DIV/0!", 5, 7), "(err-lit #DIV/0!)");
}

// ---------------------------------------------------------------------------
// Direct-construction tests for kinds without parser sugar
// ---------------------------------------------------------------------------

TEST(AstShift, StructuredRefUnchanged) {
  Arena a;
  // Tables[col] is a structured ref; no relative shift applies.
  AstNode* node = make_structured_ref(a, "Tables", "col", StructuredRefModifier::None);
  ASSERT_NE(node, nullptr);
  const AstNode* shifted = shift_relative_refs(*node, a, 5, 5);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(dump_sexpr(*shifted), "(struct-ref Tables col)");
}

TEST(AstShift, NameRefUnchanged) {
  Arena a;
  AstNode* node = make_name_ref(a, "MyName");
  ASSERT_NE(node, nullptr);
  const AstNode* shifted = shift_relative_refs(*node, a, 5, 5);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(dump_sexpr(*shifted), "(name MyName)");
}

TEST(AstShift, ExternalRefShiftsCellComponent) {
  Arena a;
  Reference cell{};
  cell.col = 0;
  cell.row = 0;
  cell.col_abs = false;
  cell.row_abs = false;
  AstNode* node = make_external_ref(a, 1, "Sheet1", cell);
  ASSERT_NE(node, nullptr);
  const AstNode* shifted = shift_relative_refs(*node, a, 2, 3);
  ASSERT_NE(shifted, nullptr);
  // The dumper format here is "(ext-ref [book] sheet ref)" — exact
  // format depends on dump_sexpr; verify the shift took effect by
  // checking the cell component appears as D3.
  const std::string out = dump_sexpr(*shifted);
  EXPECT_NE(out.find("D3"), std::string::npos) << out;
}

TEST(AstShift, ExternalRefAbsoluteCellPreserved) {
  Arena a;
  Reference cell{};
  cell.col = 0;
  cell.row = 0;
  cell.col_abs = true;
  cell.row_abs = true;
  AstNode* node = make_external_ref(a, 1, "Sheet1", cell);
  const AstNode* shifted = shift_relative_refs(*node, a, 5, 5);
  ASSERT_NE(shifted, nullptr);
  EXPECT_NE(dump_sexpr(*shifted).find("$A$1"), std::string::npos);
}

// ---------------------------------------------------------------------------
// Arena lifetime: shifted AST must outlive parser
// ---------------------------------------------------------------------------

TEST(AstShift, ShiftedTreeUsesPassedArena) {
  // Parsing and shifting in the same arena is the common path; verify
  // the result remains valid after the parser goes out of scope.
  Arena arena;
  const AstNode* shifted = nullptr;
  {
    Parser parser("A1+B2", arena);
    const AstNode* root = parser.parse();
    ASSERT_NE(root, nullptr);
    shifted = shift_relative_refs(*root, arena, 1, 0);
  }
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(dump_sexpr(*shifted), "(binary + (ref A2) (ref B3))");
}

}  // namespace
}  // namespace parser
}  // namespace formulon
