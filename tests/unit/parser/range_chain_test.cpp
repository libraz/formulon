//
// A `:` chain keeps every endpoint in the order it was written, and which
// adjacent pairs are the two halves of one whole-column or whole-row token
// (`A:C`). Every emitter therefore reproduces the typed spelling --
// `H:H:E:E` and `A:C:E:G` stay as written while `A:C` stays compact -- and
// a shift moves each endpoint.

#include <cstdint>
#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_dump.h"
#include "parser/ast_format.h"
#include "parser/ast_format_r1c1.h"
#include "parser/ast_shift.h"
#include "parser/parser.h"
#include "parser/ref_transforms.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

const AstNode* ParseClean(std::string_view src, Arena& arena) {
  Parser p(src, arena);
  const AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << src;
  EXPECT_TRUE(p.errors().empty()) << src;
  return root;
}

constexpr const char* kSpellings[] = {
    "SUM(H:H:E:E)",
    "SUM(A:C:E:G)",
    "SUM(A:C)",
    "SUM(A1:C3:E5)",
    "SUM(A1:B2:C3:D4)",
    "SUM(A:C:1:3)",
    "SUM(A:C:E1)",
    "SUM(1:3:E5)",
    "SUM(A:A:C:C)",
    "SUM(H:H)",
    "A1:B2:C3",
    "SUM($A:C:E:$G)",
    "SUM(A1:B2)",
    "SUM((A1:B2):C3)",
    "SUM(A1:(B2:C3))",
    "SUM(Sheet1!A1:B2:C3)",
    "SUM(OFFSET(A1,1,1):B2:C3)",
    "SUM(H:H:E1)",
    "A1:B2:.C3",
    "SUM(Sheet1!A:C:E:E)",
};

TEST(RangeChain, EndpointsAndSpliceFlagsAreKept) {
  Arena arena;
  const AstNode* hh = ParseClean("=H:H:E:E", arena);
  ASSERT_NE(hh, nullptr);
  EXPECT_EQ(dump_sexpr(*hh), "(range (ref H:H) (ref E:E))");
  const AstNode* he = ParseClean("=H:E", arena);
  ASSERT_NE(he, nullptr);
  EXPECT_EQ(dump_sexpr(*he), "(range (ref H:H) spliced (ref E:E))");
  const AstNode* chain = ParseClean("=A:C:E:G", arena);
  ASSERT_NE(chain, nullptr);
  ASSERT_EQ(chain->kind(), NodeKind::RangeOp);
  ASSERT_EQ(chain->as_range_arity(), 4U);
  EXPECT_TRUE(chain->as_range_spliced(0));
  EXPECT_FALSE(chain->as_range_spliced(1));
  EXPECT_TRUE(chain->as_range_spliced(2));
}

TEST(RangeChain, FormatterKeepsTheTypedSpelling) {
  for (const char* text : kSpellings) {
    Arena arena;
    const AstNode* root = ParseClean(std::string("=") + text, arena);
    ASSERT_NE(root, nullptr);
    EXPECT_EQ(format_formula(*root), text);
  }
}

TEST(RangeChain, StorageEmitterKeepsTheSpellingNextToAPrefixedFunction) {
  const auto spell = [](std::string_view name) {
    return name == "XLOOKUP" ? "_xlfn." + std::string(name) : std::string(name);
  };
  for (const char* text :
       {"SUM(H:H:E:E)+XLOOKUP(1,A1:A2,B1:B2)", "XLOOKUP(1,A:C:E:G,A1:C3:E5)", "SUM(A:C)+XLOOKUP(1,A:A,B:B)"}) {
    Arena arena;
    const std::string src = std::string("=") + text;
    const AstNode* root = ParseClean(src, arena);
    ASSERT_NE(root, nullptr);
    std::string expected = text;
    expected.replace(expected.find("XLOOKUP"), 7, "_xlfn.XLOOKUP");
    EXPECT_EQ(format_formula_storage(*root, spell), expected);
  }
}

TEST(RangeChain, R1C1PrintsEachToken) {
  Arena arena;
  const AstNode* root = ParseClean("=SUM(A:C:E:G)+SUM(A1:C3:E5)", arena);
  ASSERT_NE(root, nullptr);
  EXPECT_EQ(format_formula_r1c1(*root, /*anchor_row=*/0U, /*anchor_col=*/0U),
            "SUM(C:C[2]:C[4]:C[6])+SUM(RC:R[2]C[2]:R[4]C[4])");
}

TEST(RangeChain, FillMovesEveryEndpoint) {
  struct Case {
    const char* before;
    const char* after;
  };
  const Case cases[] = {
      {"SUM(H:H:E:E)", "SUM(I:I:F:F)"},
      {"SUM(A:C:E:G)", "SUM(B:D:F:H)"},
      {"SUM(A1:C3:E5)", "SUM(B2:D4:F6)"},
      {"SUM(A:C:1:3)", "SUM(B:D:2:4)"},
  };
  for (const Case& c : cases) {
    Arena arena;
    const AstNode* root = ParseClean(std::string("=") + c.before, arena);
    ASSERT_NE(root, nullptr);
    const AstNode* shifted = shift_relative_refs(*root, arena, /*row_delta=*/1, /*col_delta=*/1);
    ASSERT_NE(shifted, nullptr);
    EXPECT_EQ(format_formula(*shifted), c.after) << c.before;
  }
}

TEST(RangeChain, ColumnInsertMovesTokensPastIt) {
  Arena arena;
  const AstNode* root = ParseClean("=SUM(H:H:E:E)+SUM(A:C:E:G)", arena);
  ASSERT_NE(root, nullptr);
  // One column inserted before column D (0-based 3).
  const RowColShiftTransform transform("Sheet1", RowColAxis::kCol, RowColEdit::kInsert, /*index=*/3U, /*count=*/1U,
                                       /*local_means_target=*/true);
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "SUM(I:I:F:F)+SUM(A:C:F:H)");
}

TEST(RangeChain, DeletingAnAreaTokenShrinksIt) {
  Arena arena;
  const AstNode* root = ParseClean("=SUM(A1:A3:C5)", arena);
  ASSERT_NE(root, nullptr);
  // Row 1 deleted: the area A1:A3 shrinks rather than the chain becoming #REF!.
  const RowColShiftTransform transform("Sheet1", RowColAxis::kRow, RowColEdit::kDelete, /*index=*/0U, /*count=*/1U,
                                       /*local_means_target=*/true);
  const AstNode* shifted = shift_refs(*root, arena, transform);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "SUM(A1:A2:C4)");
}

TEST(RangeChain, LongChainIsOneNode) {
  std::string text = "=A1";
  for (std::uint32_t i = 1; i < 2000U; ++i) {
    text.append(":A").append(std::to_string(i % 9U + 1U));
  }
  ASSERT_LT(text.size(), 8192U);
  Arena arena;
  const AstNode* root = ParseClean(text, arena);
  ASSERT_NE(root, nullptr);
  ASSERT_EQ(root->kind(), NodeKind::RangeOp);
  EXPECT_EQ(root->as_range_arity(), 2000U);
  EXPECT_TRUE(ast_depth_within_limit(*root, 2U));
  EXPECT_EQ("=" + format_formula(*root), text);
  const AstNode* shifted = shift_relative_refs(*root, arena, /*row_delta=*/0, /*col_delta=*/1);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(shifted->as_range_arity(), 2000U);
}

TEST(RangeChain, TokenWidthFollowsLeftToRightTokenisation) {
  Arena arena;
  const AstNode* root = ParseClean("=A1:B2:C3:D4:E5", arena);
  ASSERT_NE(root, nullptr);
  EXPECT_EQ(range_token_width(*root, 0), 2U);
  EXPECT_EQ(range_token_width(*root, 2), 2U);
  EXPECT_EQ(range_token_width(*root, 4), 1U);
  const AstNode* axes = ParseClean("=H:H:E:E", arena);
  ASSERT_NE(axes, nullptr);
  EXPECT_EQ(range_token_width(*axes, 0), 1U);
}

}  // namespace
}  // namespace parser
}  // namespace formulon
