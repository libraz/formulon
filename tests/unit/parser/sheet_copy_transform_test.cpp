//
// Tests for `SheetCopyTransform` and `IdentityTransform`, driven through the
// parse -> `shift_refs` -> format path a copied sheet's formulas take.

#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/parser.h"
#include "parser/ref_transforms.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

std::string Rewrite(std::string_view src, const RefTransform& transform) {
  Arena arena;
  Parser p(src, arena);
  const AstNode* root = p.parse();
  if (root == nullptr || !p.errors().empty()) {
    return "<parse-failed>";
  }
  const AstNode* shifted = shift_refs(*root, arena, transform);
  if (shifted == nullptr) {
    return "<arena-fail>";
  }
  return format_formula(*shifted);
}

SheetCopyTransform MakeCopy() {
  return SheetCopyTransform("Src", "Src (2)", {{"Tbl", "Tbl_3"}});
}

TEST(SheetCopyTransform, SelfQualifiedCellAndRangeMoveToCopy) {
  const SheetCopyTransform t = MakeCopy();
  EXPECT_EQ(Rewrite("=Src!A1", t), "'Src (2)'!A1");
  EXPECT_EQ(Rewrite("=SUM(Src!A1:B2)", t), "SUM('Src (2)'!A1:B2)");
  EXPECT_EQ(Rewrite("=Src!$A$1+Src!C:C", t), "'Src (2)'!$A$1+'Src (2)'!C:C");
}

TEST(SheetCopyTransform, QuotedSourceMovesToCopy) {
  const SheetCopyTransform t = MakeCopy();
  EXPECT_EQ(Rewrite("='Src'!A2", t), "'Src (2)'!A2");
}

TEST(SheetCopyTransform, UnqualifiedReferenceUntouched) {
  const SheetCopyTransform t = MakeCopy();
  EXPECT_EQ(Rewrite("=A1+SUM(B1:B3)", t), "A1+SUM(B1:B3)");
}

TEST(SheetCopyTransform, OtherSheetUntouched) {
  const SheetCopyTransform t = MakeCopy();
  EXPECT_EQ(Rewrite("=Other!A1+'Src x'!B1", t), "Other!A1+'Src x'!B1");
}

TEST(SheetCopyTransform, ThreeDSpanWithSourceAsBeginIsRef) {
  const SheetCopyTransform t = MakeCopy();
  EXPECT_EQ(Rewrite("=SUM(Src:Other!A1)", t), "SUM(#REF!A1)");
}

TEST(SheetCopyTransform, ThreeDSpanWithSourceAsEndIsRef) {
  const SheetCopyTransform t = MakeCopy();
  EXPECT_EQ(Rewrite("=SUM(Other:Src!A1:B2)", t), "SUM(#REF!A1:B2)");
}

TEST(SheetCopyTransform, ThreeDSpanWithSourceAsBothEndsCollapsesToCopy) {
  const SheetCopyTransform t = MakeCopy();
  EXPECT_EQ(Rewrite("=SUM(Src:Src!A1)", t), "SUM('Src (2)'!A1)");
  EXPECT_EQ(Rewrite("=SUM(Src:Src!A1:B2)", t), "SUM('Src (2)'!A1:B2)");
  const SheetCopyTransform s("S", "S (2)", {});
  EXPECT_EQ(Rewrite("=S:S!A1", s), "'S (2)'!A1");
}

TEST(SheetCopyTransform, ThreeDSpanAroundSourceUnchanged) {
  const SheetCopyTransform t = MakeCopy();
  EXPECT_EQ(Rewrite("=SUM(First:Last!A1)", t), "SUM(First:Last!A1)");
}

TEST(SheetCopyTransform, SheetScopedNameMovesToCopy) {
  const SheetCopyTransform t = MakeCopy();
  EXPECT_EQ(Rewrite("=Src!LN+Other!LN+LN", t), "'Src (2)'!LN+Other!LN+LN");
}

TEST(SheetCopyTransform, StructuredReferenceFollowsTableMap) {
  const SheetCopyTransform t = MakeCopy();
  EXPECT_EQ(Rewrite("=SUM(Tbl[V])", t), "SUM(Tbl_3[V])");
  EXPECT_EQ(Rewrite("=SUM(tbl[V])", t), "SUM(Tbl_3[V])");
  EXPECT_EQ(Rewrite("=SUM(Other[V])", t), "SUM(Other[V])");
}

TEST(SheetCopyTransform, IndirectStringLiteralUntouched) {
  const SheetCopyTransform t = MakeCopy();
  EXPECT_EQ(Rewrite("=INDIRECT(\"Src!A1\")", t), "INDIRECT(\"Src!A1\")");
}

TEST(SheetCopyTransform, SourceMatchIsCaseInsensitive) {
  const SheetCopyTransform t = MakeCopy();
  EXPECT_EQ(Rewrite("=SRC!A1+src!LN", t), "'Src (2)'!A1+'Src (2)'!LN");
  EXPECT_EQ(Rewrite("=SUM(src:SRC!A1)", t), "SUM('Src (2)'!A1)");
}

TEST(SheetCopyTransform, BareCopyNameStaysBare) {
  const SheetCopyTransform t("Data", "Data2_", {});
  EXPECT_EQ(Rewrite("=Data!A1", t), "Data2_!A1");
}

TEST(SheetCopyTransform, IdentityTransformLeavesEveryShapeUnchanged) {
  const IdentityTransform t;
  for (const std::string_view src : {"Src!A1", "SUM(Src!A1:B2)", "'Src'!A2", "A1", "Other!A1", "SUM(Src:Other!A1)",
                                     "SUM(Other:Src!A1)", "SUM(Src:Src!A1)", "SUM(First:Last!A1)", "Src!LN",
                                     "SUM(Tbl[V])", "INDIRECT(\"Src!A1\")", "Src!#REF!", "#REF!A1"}) {
    Arena arena;
    Parser p(src, arena);
    const AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << src;
    EXPECT_EQ(shift_refs(*root, arena, t), root) << src;
  }
}

}  // namespace
}  // namespace parser
}  // namespace formulon
