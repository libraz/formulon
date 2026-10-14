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

class IdentityTransform final : public RefTransform {
 public:
  std::optional<Reference> apply(const Reference& ref) const override { return ref; }
};

class SelectiveColumnTransform final : public RefTransform {
 public:
  SelectiveColumnTransform(std::uint32_t target_col, std::uint32_t replacement_col, int* calls) noexcept
      : target_col_(target_col), replacement_col_(replacement_col), calls_(calls) {}

  std::optional<Reference> apply(const Reference& ref) const override {
    ++*calls_;
    if (ref.col != target_col_) {
      return ref;
    }
    Reference out = ref;
    out.col = replacement_col_;
    return out;
  }

 private:
  std::uint32_t target_col_;
  std::uint32_t replacement_col_;
  int* calls_;
};

class SheetTransform final : public RefTransform {
 public:
  explicit SheetTransform(std::string_view sheet) noexcept : sheet_(sheet) {}

  std::optional<Reference> apply(const Reference& ref) const override {
    Reference out = ref;
    out.sheet = sheet_;
    return out;
  }

 private:
  std::string_view sheet_;
};

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

TEST(ShiftRelativeRefs, DepthCheckPreservesLargerCallerLimit) {
  Arena arena;
  const AstNode* root = MakeTooDeepAst(arena);
  ASSERT_NE(root, nullptr);
  EXPECT_TRUE(ast_depth_within_limit(*root, kMaxFormulaAstDepth + 1));
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

TEST(ShiftRefs, IdentityWalkDoesNotAllocateVariadicChildren) {
  IdentityTransform transform;

  {
    Arena arena;
    AstNode* first = make_literal(arena, Value::number(1.0));
    AstNode* second = make_literal(arena, Value::number(2.0));
    const AstNode* args[] = {first, second};
    AstNode* call = make_call(arena, "SUM", args, 2);
    ASSERT_NE(call, nullptr);
    const std::size_t before = arena.bytes_used();
    EXPECT_EQ(shift_refs(*call, arena, transform), call);
    EXPECT_EQ(arena.bytes_used(), before);
  }
  {
    Arena arena;
    AstNode* first = make_literal(arena, Value::number(1.0));
    AstNode* second = make_literal(arena, Value::number(2.0));
    const AstNode* children[] = {first, second};
    AstNode* union_node = make_union_op(arena, children, 2);
    ASSERT_NE(union_node, nullptr);
    const std::size_t before = arena.bytes_used();
    EXPECT_EQ(shift_refs(*union_node, arena, transform), union_node);
    EXPECT_EQ(arena.bytes_used(), before);
  }
  {
    Arena arena;
    AstNode* first = make_literal(arena, Value::number(1.0));
    AstNode* second = make_literal(arena, Value::number(2.0));
    const AstNode* elements[] = {first, second};
    AstNode* array = make_array_literal(arena, 1, 2, elements);
    ASSERT_NE(array, nullptr);
    const std::size_t before = arena.bytes_used();
    EXPECT_EQ(shift_refs(*array, arena, transform), array);
    EXPECT_EQ(arena.bytes_used(), before);
  }
  {
    Arena arena;
    AstNode* binding_expr = make_literal(arena, Value::number(1.0));
    AstNode* body = make_literal(arena, Value::number(2.0));
    const std::string_view names[] = {"x"};
    const AstNode* exprs[] = {binding_expr};
    AstNode* let_node = make_let_binding(arena, names, exprs, 1, body);
    ASSERT_NE(let_node, nullptr);
    const std::size_t before = arena.bytes_used();
    EXPECT_EQ(shift_refs(*let_node, arena, transform), let_node);
    EXPECT_EQ(arena.bytes_used(), before);
  }
  {
    Arena arena;
    AstNode* callee = make_name_ref(arena, "fn");
    AstNode* arg = make_literal(arena, Value::number(1.0));
    const AstNode* args[] = {arg};
    AstNode* lambda_call = make_lambda_call(arena, callee, args, 1);
    ASSERT_NE(lambda_call, nullptr);
    const std::size_t before = arena.bytes_used();
    EXPECT_EQ(shift_refs(*lambda_call, arena, transform), lambda_call);
    EXPECT_EQ(arena.bytes_used(), before);
  }
}

TEST(ShiftRefs, ChangedMiddleChildPreservesPrefixAndAppliesOnce) {
  Arena arena;
  AstNode* first = make_ref(arena, Reference{});
  Reference middle_ref;
  middle_ref.col = 1;
  AstNode* middle = make_ref(arena, middle_ref);
  Reference last_ref;
  last_ref.col = 2;
  AstNode* last = make_ref(arena, last_ref);
  const AstNode* args[] = {first, middle, last};
  AstNode* call = make_call(arena, "SUM", args, 3);
  ASSERT_NE(call, nullptr);

  int calls = 0;
  SelectiveColumnTransform transform(1, 9, &calls);
  const AstNode* shifted = shift_refs(*call, arena, transform);
  ASSERT_NE(shifted, nullptr);
  ASSERT_EQ(shifted->kind(), NodeKind::Call);
  EXPECT_EQ(calls, 3);
  EXPECT_EQ(&shifted->as_call_arg(0), first);
  EXPECT_NE(&shifted->as_call_arg(1), middle);
  EXPECT_EQ(shifted->as_call_arg(1).as_ref().col, 9u);
  EXPECT_EQ(&shifted->as_call_arg(2), last);
}

TEST(ShiftRefs, PropagatesInternFailureFromChangedReference) {
  Arena source_arena;
  Reference source_ref;
  source_ref.sheet = "Sheet1";
  AstNode* root = make_ref(source_arena, source_ref);
  ASSERT_NE(root, nullptr);

  const std::string long_sheet(512, 'x');
  Arena destination_arena(256, 256);
  SheetTransform transform(long_sheet);
  EXPECT_EQ(shift_refs(*root, destination_arena, transform), nullptr);
}

TEST(ShiftRefs, LetBodyOnlyChangePreservesBindingExpression) {
  Arena arena;
  AstNode* binding_expr = make_literal(arena, Value::number(1.0));
  Reference body_ref;
  body_ref.col = 1;
  AstNode* body = make_ref(arena, body_ref);
  const std::string_view names[] = {"x"};
  const AstNode* exprs[] = {binding_expr};
  AstNode* let_node = make_let_binding(arena, names, exprs, 1, body);
  ASSERT_NE(let_node, nullptr);

  int calls = 0;
  SelectiveColumnTransform transform(1, 9, &calls);
  const AstNode* shifted = shift_refs(*let_node, arena, transform);
  ASSERT_NE(shifted, nullptr);
  ASSERT_EQ(shifted->kind(), NodeKind::LetBinding);
  EXPECT_EQ(calls, 1);
  EXPECT_EQ(&shifted->as_let_binding_expr(0), binding_expr);
  EXPECT_NE(&shifted->as_let_body(), body);
  EXPECT_EQ(shifted->as_let_body().as_ref().col, 9u);
}

TEST(ShiftRefs, LambdaCalleeChangePreservesArguments) {
  Arena arena;
  Reference callee_ref;
  callee_ref.col = 1;
  AstNode* callee = make_ref(arena, callee_ref);
  Reference arg_ref;
  arg_ref.col = 0;
  AstNode* arg = make_ref(arena, arg_ref);
  const AstNode* args[] = {arg};
  AstNode* lambda_call = make_lambda_call(arena, callee, args, 1);
  ASSERT_NE(lambda_call, nullptr);

  int calls = 0;
  SelectiveColumnTransform transform(1, 9, &calls);
  const AstNode* shifted = shift_refs(*lambda_call, arena, transform);
  ASSERT_NE(shifted, nullptr);
  ASSERT_EQ(shifted->kind(), NodeKind::LambdaCall);
  EXPECT_EQ(calls, 2);
  EXPECT_NE(&shifted->as_lambda_call_callee(), callee);
  EXPECT_EQ(shifted->as_lambda_call_callee().as_ref().col, 9u);
  EXPECT_EQ(&shifted->as_lambda_call_arg(0), arg);
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

TEST(ShiftRefsWithSheetRename, RenamesSheetQualifiedNames) {
  Arena arena;
  const AstNode* root = ParseOrNull("=Sheet1!Rate+Other!Rate+Rate", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("sheet1", "My Sheet");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "'My Sheet'!Rate+Other!Rate+Rate");
}

TEST(ShiftRefsWithSheetRename, RenamesSheetQualifiedNameCalls) {
  Arena arena;
  const AstNode* root = ParseOrNull("=Sheet1!Fn(Sheet1!A1)+Fn(1)", arena);
  ASSERT_NE(root, nullptr);
  SheetRenameTransform transform("Sheet1", "My Sheet");
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "'My Sheet'!Fn('My Sheet'!A1)+Fn(1)");
}

TEST(ShiftRefsWithSheetRemoval, SheetQualifiedNameOnRemovedSheetBecomesRef) {
  EXPECT_EQ(ParseSheetRemovalFormula("=Sheet2!Rate+Sheet1!Rate+Rate", {"Sheet1", "Sheet2"}, 1U),
            "#REF!+Sheet1!Rate+Rate");
}

// The whole call collapses: `#REF!(3)` would not parse back.
TEST(ShiftRefsWithSheetRemoval, SheetQualifiedNameCallOnRemovedSheetBecomesRef) {
  EXPECT_EQ(ParseSheetRemovalFormula("=Sheet2!Fn(3)+Sheet1!Fn(A1)", {"Sheet1", "Sheet2"}, 1U), "#REF!+Sheet1!Fn(A1)");
}

TEST(ShiftRefsWithRowColShift, SheetQualifiedNameIsUntouched) {
  EXPECT_EQ(ParseRowColShiftFormula("=Sheet1!Rate+A2", RowColAxis::kRow, RowColEdit::kInsert, 0U, 1U),
            "Sheet1!Rate+A3");
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
