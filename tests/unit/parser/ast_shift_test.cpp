//
// Tests for `shift_refs` (generic walker), `shift_relative_refs` (the
// historical relative-shift wrapper), and the integration with the
// sheet-structure transforms.

#include "parser/ast_shift.h"

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_dump.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "parser/ref_transforms.h"
#include "parser/reference.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

// Helper: parse `src`, return root and accumulate any errors. The arena
// must outlive the returned pointer.
const AstNode* ParseOrNull(std::string_view src, Arena& arena) {
  Parser p(src, arena);
  AstNode* root = p.parse();
  if (!p.errors().empty()) {
    return nullptr;
  }
  return root;
}

const AstNode* MakeTooDeepAst(Arena& arena) {
  AstNode* node = make_literal(arena, Value::number(1));
  for (std::uint32_t depth = 1; depth <= kMaxFormulaAstDepth; ++depth) {
    node = make_unary_op(arena, UnaryOp::Plus, node);
  }
  return node;
}

std::string ParseShiftRelativeDump(std::string_view src, std::int32_t row_delta, std::int32_t col_delta) {
  Arena arena;
  const AstNode* root = ParseOrNull(src, arena);
  if (root == nullptr) {
    return "<parse-failed>";
  }
  const AstNode* shifted = shift_relative_refs(*root, arena, row_delta, col_delta);
  if (shifted == nullptr) {
    return "<arena-fail>";
  }
  return dump_sexpr(*shifted);
}

std::string ParseRowColShiftFormula(std::string_view src, RowColAxis axis, RowColEdit edit, std::uint32_t index,
                                    std::uint32_t count, std::string_view target_sheet = "Sheet1",
                                    bool local_means_target = true) {
  Arena arena;
  const AstNode* root = ParseOrNull(src, arena);
  if (root == nullptr) {
    return "<parse-failed>";
  }
  RowColShiftTransform transform(target_sheet, axis, edit, index, count, local_means_target);
  const AstNode* shifted = shift_refs(*root, arena, transform);
  if (shifted == nullptr) {
    return "<arena-fail>";
  }
  return format_formula(*shifted);
}

std::string ParseSheetRemovalFormula(std::string_view src, const std::vector<std::string_view>& order,
                                     std::uint32_t removed_index) {
  Arena arena;
  const AstNode* root = ParseOrNull(src, arena);
  if (root == nullptr) {
    return "<parse-failed>";
  }
  SheetRemovalTransform transform(order, removed_index);
  const AstNode* shifted = shift_refs(*root, arena, transform);
  if (shifted == nullptr) {
    return "<arena-fail>";
  }
  return format_formula(*shifted);
}

// ---------------------------------------------------------------------------
// shift_relative_refs (legacy wrapper)
// ---------------------------------------------------------------------------

TEST(ShiftRelativeRefs, IdentityZeroDelta) {
  EXPECT_EQ(ParseShiftRelativeDump("=A1", 0, 0), "(ref A1)");
}

TEST(ShiftRelativeRefs, RejectsAstDeeperThanSharedLimit) {
  Arena arena;
  const AstNode* root = MakeTooDeepAst(arena);
  ASSERT_NE(root, nullptr);
  const AstNode* shifted = shift_relative_refs(*root, arena, 0, 0);
  ASSERT_NE(shifted, nullptr);
  ASSERT_EQ(shifted->kind(), NodeKind::ErrorLiteral);
  EXPECT_EQ(shifted->as_error_literal(), ErrorCode::Ref);
}

TEST(ShiftRelativeRefs, RowDeltaShiftsRelativeRow) {
  EXPECT_EQ(ParseShiftRelativeDump("=A1", 2, 0), "(ref A3)");
}

TEST(ShiftRelativeRefs, ColDeltaShiftsRelativeCol) {
  EXPECT_EQ(ParseShiftRelativeDump("=A1", 0, 2), "(ref C1)");
}

TEST(ShiftRelativeRefs, WrittenOutSpillAnchorShifts) {
  EXPECT_EQ(ParseShiftRelativeDump("=A1#", 2, 0), "(spill-ref A3#)");
}

TEST(ShiftRelativeRefs, ComputedSpillAnchorShiftsInsideTheExpression) {
  // The operator itself has no reference to move; the shift lands on the
  // refs the anchor expression reads.
  EXPECT_EQ(ParseShiftRelativeDump("=OFFSET(A1,1,0)#", 2, 0), "(spill-ref (call OFFSET (ref A3) (num 1) (num 0))#)");
}

TEST(ShiftRelativeRefs, AbsoluteRowKept) {
  EXPECT_EQ(ParseShiftRelativeDump("=A$1", 5, 0), "(ref A$1)");
}

TEST(ShiftRelativeRefs, AbsoluteColKept) {
  EXPECT_EQ(ParseShiftRelativeDump("=$A1", 0, 5), "(ref $A1)");
}

TEST(ShiftRelativeRefs, BothAbsoluteUnchanged) {
  EXPECT_EQ(ParseShiftRelativeDump("=$A$1", 5, 5), "(ref $A$1)");
}

TEST(ShiftRelativeRefs, OutOfBoundsCollapsesToRef) {
  // A1 with row_delta=-1 produces row=-1 → out of bounds → #REF!.
  EXPECT_EQ(ParseShiftRelativeDump("=A1", -1, 0), "(err-lit #REF!)");
}

TEST(ShiftRelativeRefs, OutOfBoundsColumnCollapsesToRef) {
  EXPECT_EQ(ParseShiftRelativeDump("=A1", 0, -1), "(err-lit #REF!)");
}

TEST(ShiftRelativeRefs, ShiftsBinaryOperands) {
  EXPECT_EQ(ParseShiftRelativeDump("=A1+B2", 1, 1), "(binary + (ref B2) (ref C3))");
}

TEST(ShiftRelativeRefs, RangeOpShiftsBothEndpoints) {
  EXPECT_EQ(ParseShiftRelativeDump("=A1:B2", 1, 1), "(range (ref B2) (ref C3))");
}

TEST(ShiftRelativeRefs, FullColumnUnchangedByRowDelta) {
  EXPECT_EQ(ParseShiftRelativeDump("=A:A", 5, 0), "(ref A:A)");
}

TEST(ShiftRelativeRefs, FullColumnShiftedByColDelta) {
  EXPECT_EQ(ParseShiftRelativeDump("=A:A", 0, 1), "(ref B:B)");
}

TEST(ShiftRelativeRefs, FullRowShiftedByRowDelta) {
  EXPECT_EQ(ParseShiftRelativeDump("=1:1", 1, 0), "(ref 2:2)");
}

TEST(ShiftRelativeRefs, NameRefUntouched) {
  EXPECT_EQ(ParseShiftRelativeDump("=foo", 1, 1), "(name foo)");
}

TEST(ShiftRelativeRefs, IdentityWalkPreservesPointer) {
  Arena arena;
  const AstNode* root = ParseOrNull("=$A$1+$B$2", arena);
  ASSERT_NE(root, nullptr);
  const AstNode* shifted = shift_relative_refs(*root, arena, 5, 5);
  EXPECT_EQ(shifted, root) << "absolute-only refs should round-trip without allocation";
}

// ---------------------------------------------------------------------------
// Range-level row/column transforms
// ---------------------------------------------------------------------------

TEST(ShiftRefsWithRowCol, RowDeleteShrinksFirstEndpoint) {
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(A1:A3)", RowColAxis::kRow, RowColEdit::kDelete, 0, 1), "SUM(A1:A2)");
}

TEST(ShiftRefsWithRowCol, RowDeleteShrinksLastEndpoint) {
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(A1:A3)", RowColAxis::kRow, RowColEdit::kDelete, 2, 1), "SUM(A1:A2)");
}

TEST(ShiftRefsWithRowCol, RowDeleteShrinksMiddleAndPreservesAbsoluteFlags) {
  EXPECT_EQ(ParseRowColShiftFormula("=SUM($A$1:$A$5)", RowColAxis::kRow, RowColEdit::kDelete, 1, 2), "SUM($A$1:$A$3)");
}

TEST(ShiftRefsWithRowCol, RowDeleteCollapsesSingletonReference) {
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(A2:A2)", RowColAxis::kRow, RowColEdit::kDelete, 1, 1), "SUM(#REF!)");
}

TEST(ShiftRefsWithRowCol, RowDeleteCollapsesFullyDeletedMultiCellRange) {
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(A2:A4)", RowColAxis::kRow, RowColEdit::kDelete, 1, 3), "SUM(#REF!)");
}

TEST(ShiftRefsWithRowCol, ColDeleteShrinksRange) {
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(A1:D1)", RowColAxis::kCol, RowColEdit::kDelete, 1, 1), "SUM(A1:C1)");
}

TEST(ShiftRefsWithRowCol, WholeRowAndColumnRangesShrink) {
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(1:4)", RowColAxis::kRow, RowColEdit::kDelete, 1, 1), "SUM(1:3)");
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(A:D)", RowColAxis::kCol, RowColEdit::kDelete, 1, 1), "SUM(A:C)");
}

TEST(ShiftRefsWithRowCol, QualifiedRangeInheritsSheetForBothEndpoints) {
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(Sheet1!A1:A3)", RowColAxis::kRow, RowColEdit::kDelete, 0, 1, "Sheet1", false),
            "SUM(Sheet1!A1:A2)");
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(Sheet2!A1:A3)", RowColAxis::kRow, RowColEdit::kDelete, 0, 1, "Sheet1", false),
            "SUM(Sheet2!A1:A3)");
}

TEST(ShiftRefsWithRowCol, StructuralEditsPreserveThreeDCoordinates) {
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(Sheet1:Sheet2!A1:A3)", RowColAxis::kRow, RowColEdit::kDelete, 0, 1, "", true),
            "SUM(Sheet1:Sheet2!A1:A3)");
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(Sheet1:Sheet2!A3)", RowColAxis::kRow, RowColEdit::kDelete, 0, 1, "", true),
            "SUM(Sheet1:Sheet2!A3)");
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(Sheet1:Sheet2!A1:C1)", RowColAxis::kCol, RowColEdit::kDelete, 0, 1, "", true),
            "SUM(Sheet1:Sheet2!A1:C1)");
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(Sheet1:Sheet2!A1:A3)", RowColAxis::kRow, RowColEdit::kInsert, 0, 1, "", true),
            "SUM(Sheet1:Sheet2!A1:A3)");
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(Sheet1:Sheet2!A1:A3)", RowColAxis::kRow, RowColEdit::kDelete, 0, 1, "Sheet1",
                                    false),
            "SUM(Sheet1:Sheet2!A1:A3)");
}

TEST(ShiftRefsWithRowCol, CrossSheetQualifiedRangesOnlyRewriteTargetSheet) {
  EXPECT_EQ(ParseRowColShiftFormula("=SUM(Sheet1!A1:A3)+SUM(Sheet2!A1:A3)", RowColAxis::kRow, RowColEdit::kDelete, 0, 1,
                                    "Sheet1", false),
            "SUM(Sheet1!A1:A2)+SUM(Sheet2!A1:A3)");
}

// ---------------------------------------------------------------------------
// shift_refs with SheetRenameTransform
// ---------------------------------------------------------------------------

TEST(ShiftRefsWithSheetRename, RenamesQualifiedRefs) {
  Arena arena;
  const AstNode* root = ParseOrNull("=Sheet1!A1+Sheet1!$B$2", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("Sheet1", "Renamed");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "Renamed!A1+Renamed!$B$2");
}

TEST(ShiftRefsWithSheetRename, RenameMatchIsCaseInsensitive) {
  Arena arena;
  const AstNode* root = ParseOrNull("=sheet1!A1", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("Sheet1", "Renamed");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "Renamed!A1");
}

TEST(ShiftRefsWithSheetRename, QuotedSourceSheetRenamesToBareTarget) {
  Arena arena;
  const AstNode* root = ParseOrNull("='Sheet 1'!A1", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("Sheet 1", "Sheet1");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "Sheet1!A1");
}

TEST(ShiftRefsWithSheetRename, BareSourceSheetRenamesToQuotedTarget) {
  Arena arena;
  const AstNode* root = ParseOrNull("=Sheet1!A1", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("Sheet1", "New Sheet");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "'New Sheet'!A1");
}

TEST(ShiftRefsWithSheetRename, NonMatchingSheetIsUnchanged) {
  Arena arena;
  const AstNode* root = ParseOrNull("=Sheet2!A1", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("Sheet1", "Renamed");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  EXPECT_EQ(shifted, root) << "unrelated sheet should not allocate";
}

TEST(ShiftRefsWithSheetRename, LocalRefsUnchanged) {
  Arena arena;
  const AstNode* root = ParseOrNull("=A1+B2", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("Sheet1", "Renamed");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  EXPECT_EQ(shifted, root) << "local references should not allocate";
}

TEST(ShiftRefsWithSheetRename, RangeWithSheetRenamed) {
  Arena arena;
  const AstNode* root = ParseOrNull("=Sheet1!A1:Sheet1!B2", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("Sheet1", "Renamed");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "Renamed!A1:Renamed!B2");
}

TEST(ShiftRefsWithSheetRename, RenamesBoth3DSpanEndpoints) {
  Arena arena;
  const AstNode* root = ParseOrNull("=SUM(Sheet1:Sheet2!A1:B2)", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("Sheet1", "Renamed");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "SUM(Renamed:Sheet2!A1:B2)");
}

TEST(ShiftRefsWithSheetRename, RenamesEitherBare3DSpanEndpoint) {
  Arena arena;
  const AstNode* root = ParseOrNull("=SUM(Sheet1:Sheet2!A1)", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("Sheet2", "Renamed");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "SUM(Sheet1:Renamed!A1)");
}

TEST(ShiftRefsWithSheetRename, RenamesEitherQuoted3DSpanEndpoint) {
  Arena arena;
  const AstNode* root = ParseOrNull("=SUM('Data:S2'!A1)", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("S2", "Summary");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "SUM(Data:Summary!A1)");
}

TEST(ShiftRefsWithSheetRename, Quoted3DSpanUsesQuotedRenamedEndpoint) {
  Arena arena;
  const AstNode* root = ParseOrNull("=SUM('Data:S2'!A1)", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("S2", "New Sheet");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "SUM('Data:New Sheet'!A1)");
}

TEST(ShiftRefsWithSheetRename, Quoted3DSpanRenamesBeginEndpoint) {
  Arena arena;
  const AstNode* root = ParseOrNull("=SUM('S1:Data'!A1)", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("S1", "New Sheet");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "SUM('New Sheet:Data'!A1)");
}

TEST(ShiftRefsWithSheetRename, Unrelated3DSpanPreservesIdentity) {
  Arena arena;
  const AstNode* root = ParseOrNull("=SUM(Sheet1:Sheet2!A1)", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("Other", "Renamed");
  EXPECT_EQ(shift_refs(*root, arena, transform), root);
}

TEST(ShiftRefs, RelativeShiftUpdates3DRangeTail) {
  Arena arena;
  const AstNode* root = ParseOrNull("=Sheet1:Sheet2!A1:B2", arena);
  ASSERT_NE(root, nullptr);
  const AstNode* shifted = shift_relative_refs(*root, arena, 2, 3);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "Sheet1:Sheet2!D3:E4");
}

// ---------------------------------------------------------------------------
// shift_refs with SheetRemovalTransform
// ---------------------------------------------------------------------------

TEST(ShiftRefsWithSheetRemoval, RemovesNormalQualifiedReference) {
  const std::vector<std::string_view> order = {"Sheet1", "Sheet2", "Sheet3"};
  EXPECT_EQ(ParseSheetRemovalFormula("=Sheet2!A1", order, /*removed_index=*/1), "#REF!");
  EXPECT_EQ(ParseSheetRemovalFormula("=SUM(Sheet2!A1:B2)+1", order, /*removed_index=*/1), "SUM(#REF!)+1");
  EXPECT_EQ(ParseSheetRemovalFormula("=Sheet1!A1", order, /*removed_index=*/1), "Sheet1!A1");
  EXPECT_EQ(ParseSheetRemovalFormula("=A1", order, /*removed_index=*/1), "A1");
}

TEST(ShiftRefsWithSheetRemoval, MiddleSpanKeepsEndpoints) {
  const std::vector<std::string_view> order = {"Sheet1", "Sheet2", "Sheet3"};
  EXPECT_EQ(ParseSheetRemovalFormula("=SUM(Sheet1:Sheet3!A1)", order, /*removed_index=*/1), "SUM(Sheet1:Sheet3!A1)");
}

TEST(ShiftRefsWithSheetRemoval, MiddleSpanPreservesIdentity) {
  Arena arena;
  const AstNode* root = ParseOrNull("=SUM(Sheet1:Sheet3!A1)", arena);
  ASSERT_NE(root, nullptr);
  const std::vector<std::string_view> order = {"Sheet1", "Sheet2", "Sheet3"};
  SheetRemovalTransform transform(order, /*removed_index=*/1);
  EXPECT_EQ(shift_refs(*root, arena, transform), root);
}

TEST(ShiftRefsWithSheetRemoval, BeginAndEndSpansMoveInward) {
  const std::vector<std::string_view> order = {"Sheet1", "Sheet2", "Sheet3"};
  EXPECT_EQ(ParseSheetRemovalFormula("=SUM(Sheet1:Sheet3!A1)", order, /*removed_index=*/0), "SUM(Sheet2:Sheet3!A1)");
  EXPECT_EQ(ParseSheetRemovalFormula("=SUM(Sheet1:Sheet3!A1)", order, /*removed_index=*/2), "SUM(Sheet1:Sheet2!A1)");
}

TEST(ShiftRefsWithSheetRemoval, ReverseSpanMovesInward) {
  const std::vector<std::string_view> order = {"Sheet1", "Sheet2", "Sheet3"};
  EXPECT_EQ(ParseSheetRemovalFormula("=SUM(Sheet3:Sheet1!A1)", order, /*removed_index=*/2), "SUM(Sheet2:Sheet1!A1)");
  EXPECT_EQ(ParseSheetRemovalFormula("=SUM(Sheet3:Sheet1!A1)", order, /*removed_index=*/0), "SUM(Sheet3:Sheet2!A1)");
}

TEST(ShiftRefsWithSheetRemoval, DegenerateAndUnresolvedSpansCollapse) {
  const std::vector<std::string_view> order = {"Sheet1", "Sheet2", "Sheet3"};
  EXPECT_EQ(ParseSheetRemovalFormula("=SUM(Sheet2:Sheet2!A1)", order, /*removed_index=*/1), "SUM(#REF!)");
  EXPECT_EQ(ParseSheetRemovalFormula("=SUM(Missing:Sheet3!A1)", order, /*removed_index=*/1), "SUM(Missing:Sheet3!A1)");
  EXPECT_EQ(ParseSheetRemovalFormula("=SUM(Sheet2:Missing!A1)", order, /*removed_index=*/1), "SUM(#REF!)");
}

TEST(ShiftRefsWithSheetRemoval, UnresolvedSpanOutsideRemovalPreservesAstIdentity) {
  Arena arena;
  const AstNode* root = ParseOrNull("=SUM(Missing:Sheet3!A1)", arena);
  ASSERT_NE(root, nullptr);
  const std::vector<std::string_view> order = {"Sheet1", "Sheet2", "Sheet3"};
  SheetRemovalTransform transform(order, /*removed_index=*/1);
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(shifted, root);
  EXPECT_EQ(format_formula(*shifted), "SUM(Missing:Sheet3!A1)");
}

TEST(ShiftRefsWithSheetRemoval, StringLiteralIsUntouched) {
  const std::vector<std::string_view> order = {"Sheet1", "Sheet2"};
  Arena arena;
  const AstNode* literal = ParseOrNull("=\"Sheet2!A1\"", arena);
  ASSERT_NE(literal, nullptr);
  SheetRemovalTransform transform(order, /*removed_index=*/1);
  EXPECT_EQ(shift_refs(*literal, arena, transform), literal);
}

// ---------------------------------------------------------------------------
// Direct RefTransform usage with a custom subclass
// ---------------------------------------------------------------------------

namespace {

// Test fixture transform: rewrites every Reference's column to 5
// regardless of input. Demonstrates that a one-line custom RefTransform
// integrates cleanly.
class FixedColTransform final : public RefTransform {
 public:
  std::optional<Reference> apply(const Reference& ref) const override {
    Reference out = ref;
    out.col = 5;
    return out;
  }
};

}  // namespace

TEST(ShiftRefsCustom, FixedColTransformRewritesEveryRef) {
  Arena arena;
  const AstNode* root = ParseOrNull("=A1+B2", arena);
  ASSERT_NE(root, nullptr);
  FixedColTransform transform;
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(dump_sexpr(*shifted), "(binary + (ref F1) (ref F2))");
}

TEST(ShiftRefsCustom, DefaultRangeHookTransformsBothEndpoints) {
  Arena arena;
  const AstNode* root = ParseOrNull("=A1:B2", arena);
  ASSERT_NE(root, nullptr);
  FixedColTransform transform;
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "F1:F2");
}

}  // namespace
}  // namespace parser
}  // namespace formulon
