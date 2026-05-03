// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tests for `shift_refs` (generic walker), `shift_relative_refs` (the
// historical relative-shift wrapper), and the integration with
// `SheetRenameTransform`.

#include "parser/ast_shift.h"

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

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

// ---------------------------------------------------------------------------
// shift_relative_refs (legacy wrapper)
// ---------------------------------------------------------------------------

TEST(ShiftRelativeRefs, IdentityZeroDelta) {
  EXPECT_EQ(ParseShiftRelativeDump("=A1", 0, 0), "(ref A1)");
}

TEST(ShiftRelativeRefs, RowDeltaShiftsRelativeRow) {
  EXPECT_EQ(ParseShiftRelativeDump("=A1", 2, 0), "(ref A3)");
}

TEST(ShiftRelativeRefs, ColDeltaShiftsRelativeCol) {
  EXPECT_EQ(ParseShiftRelativeDump("=A1", 0, 2), "(ref C1)");
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

TEST(ShiftRefsWithSheetRename, ExternalRefSheetRenamed) {
  // External refs are not parser-produceable; build the AST manually.
  Arena arena;
  Reference cell;
  cell.col = 0;
  cell.row = 0;
  AstNode* ext = make_external_ref(arena, 1, "Sheet1", cell);
  ASSERT_NE(ext, nullptr);
  SheetRenameTransform transform("Sheet1", "Renamed");
  const AstNode* shifted = shift_refs(*ext, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "[1]Renamed!A1");
}

TEST(ShiftRefsWithSheetRename, ExternalRefUnrelatedSheetUnchanged) {
  Arena arena;
  Reference cell;
  cell.col = 0;
  cell.row = 0;
  AstNode* ext = make_external_ref(arena, 1, "OtherSheet", cell);
  ASSERT_NE(ext, nullptr);
  SheetRenameTransform transform("Sheet1", "Renamed");
  const AstNode* shifted = shift_refs(*ext, arena, transform);
  EXPECT_EQ(shifted, ext) << "non-matching external sheet should be a no-op";
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

}  // namespace
}  // namespace parser
}  // namespace formulon
