// Dependency extractor tests grouped by dependency source.

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "defined_name.h"
#include "dep_extractor_test_helpers.h"
#include "eval/dep_extractor.h"
#include "table.h"
#include "utils/resource_budget.h"
#include "workbook.h"

namespace formulon::eval {
namespace {
using namespace dep_extractor_test_helpers;

TEST(DepExtractor, NameRefWorkbookScopedSingleCell) {
  Workbook wb = Workbook::create();
  // MyName -> =A1, workbook-scoped (local_sheet_id = -1).
  std::vector<DefinedName> names;
  names.push_back(DefinedName{"MyName", "=A1", -1, false, ""});
  wb.set_defined_names(std::move(names));

  Arena arena;
  const parser::AstNode* root = ParseFormula("MyName+1", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, /*current_sheet_id=*/0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  ASSERT_EQ(deps.cell_deps.size(), 1u);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{0U, 0U, 0U}));
}

TEST(DepExtractor, NameRefWorkbookScopedRangeFlattens) {
  Workbook wb = Workbook::create();
  // MyRange -> =A1:B2, workbook-scoped. SUM(MyRange) should flatten to
  // {A1, A2, B1, B2}.
  std::vector<DefinedName> names;
  names.push_back(DefinedName{"MyRange", "=A1:B2", -1, false, ""});
  wb.set_defined_names(std::move(names));

  Arena arena;
  const parser::AstNode* root = ParseFormula("SUM(MyRange)", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  std::vector<CellNodeId> expected = {
      CellNodeId{0U, 0U, 0U},  // A1
      CellNodeId{0U, 0U, 1U},  // B1
      CellNodeId{0U, 1U, 0U},  // A2
      CellNodeId{0U, 1U, 1U},  // B2
  };
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, NameRefIndirectionResolves) {
  Workbook wb = Workbook::create();
  // Name1 -> =Name2, Name2 -> =A1. =Name1 must surface A1 as a dep.
  std::vector<DefinedName> names;
  names.push_back(DefinedName{"Name1", "=Name2", -1, false, ""});
  names.push_back(DefinedName{"Name2", "=A1", -1, false, ""});
  wb.set_defined_names(std::move(names));

  Arena arena;
  const parser::AstNode* root = ParseFormula("Name1", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  ASSERT_EQ(deps.cell_deps.size(), 1u);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{0U, 0U, 0U}));
}

TEST(DepExtractor, NameRefSheetScopedBeatsWorkbookScoped) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");  // index 1
  // Foo at workbook scope -> =A1; Foo at sheet 0 -> =B1. From sheet 0 the
  // sheet-scoped definition wins (B1); from sheet 1 the workbook-scoped
  // fallback applies (A1 on sheet 1, since unqualified refs resolve to the
  // current sheet).
  std::vector<DefinedName> names;
  names.push_back(DefinedName{"Foo", "=A1", -1, false, ""});
  names.push_back(DefinedName{"Foo", "=B1", 0, false, ""});
  wb.set_defined_names(std::move(names));

  Arena arena;
  const parser::AstNode* root = ParseFormula("Foo", arena);
  ASSERT_NE(root, nullptr);

  ExtractedDeps from_sheet0 = extract_deps(*root, /*current_sheet_id=*/0U, wb);
  ASSERT_EQ(from_sheet0.cell_deps.size(), 1u);
  EXPECT_EQ(from_sheet0.cell_deps[0], (CellNodeId{0U, 0U, 1U}));  // B1 on sheet 0

  ExtractedDeps from_sheet1 = extract_deps(*root, /*current_sheet_id=*/1U, wb);
  ASSERT_EQ(from_sheet1.cell_deps.size(), 1u);
  EXPECT_EQ(from_sheet1.cell_deps[0], (CellNodeId{1U, 0U, 0U}));  // A1 on sheet 1
}

TEST(DepExtractor, NameRefCycleTerminates) {
  Workbook wb = Workbook::create();
  // Loop -> =Loop+1 — self-referential. The walker must not infinite-loop;
  // policy is to break the cycle silently (no deps, no volatility flag).
  std::vector<DefinedName> names;
  names.push_back(DefinedName{"Loop", "=Loop+1", -1, false, ""});
  wb.set_defined_names(std::move(names));

  Arena arena;
  const parser::AstNode* root = ParseFormula("Loop", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  // The assertion that matters is that the test terminates. Cycle policy:
  // silent skip on re-entry — no deps, not volatile.
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, NameRefMissingNameIsSilentSkip) {
  Workbook wb = Workbook::create();
  // No defined names registered: =MissingName must not crash and produces
  // an empty dep set.
  Arena arena;
  const parser::AstNode* root = ParseFormula("MissingName", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, NameRefVolatileBodyPropagates) {
  Workbook wb = Workbook::create();
  // RandName -> =RAND(). =RandName must propagate volatility.
  std::vector<DefinedName> names;
  names.push_back(DefinedName{"RandName", "=RAND()", -1, false, ""});
  wb.set_defined_names(std::move(names));

  Arena arena;
  const parser::AstNode* root = ParseFormula("RandName", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_TRUE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, NamedLambdaBodyCellsAndVolatilityAreExpandedOnce) {
  Workbook wb = Workbook::create();
  wb.set_defined_names({DefinedName{"Named", "LAMBDA(x,x+A1+RAND())", -1, false, ""}});
  Arena arena;
  const parser::AstNode* root = ParseFormula("Named(5)", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_TRUE(deps.is_volatile);
  ASSERT_EQ(deps.cell_deps.size(), 1U);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{0U, 0U, 0U}));  // A1
}

TEST(DepExtractor, NamedLambdaParameterShadowsDefinedName) {
  Workbook wb = Workbook::create();
  wb.set_defined_names({
      DefinedName{"x", "B1", -1, false, ""},
      DefinedName{"F", "LAMBDA(x,x+A1)", -1, false, ""},
  });
  Arena arena;
  const parser::AstNode* root = ParseFormula("F(5)", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  ASSERT_EQ(deps.cell_deps.size(), 1U);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{0U, 0U, 0U}));  // A1, not x -> B1
}

TEST(DepExtractor, NamedLambdaDoesNotInheritCallerLetScope) {
  Workbook wb = Workbook::create();
  wb.set_defined_names({
      DefinedName{"x", "B1", -1, false, ""},
      DefinedName{"F", "LAMBDA(y,y+x)", -1, false, ""},
  });
  Arena arena;
  const parser::AstNode* root = ParseFormula("LET(x,A1,F(1)+x)", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  std::vector<CellNodeId> expected = {
      CellNodeId{0U, 0U, 0U},  // LET initializer A1, read through x
      CellNodeId{0U, 0U, 1U},  // defined x -> B1 inside F
  };
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, NamedLambdaRecursiveBodyIsFinite) {
  Workbook wb = Workbook::create();
  wb.set_defined_names({DefinedName{"Fact", "LAMBDA(n,IF(n<=1,1,n*Fact(n-1)+A1))", -1, false, ""}});
  Arena arena;
  const parser::AstNode* root = ParseFormula("Fact(5)", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  ASSERT_EQ(deps.cell_deps.size(), 1U);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{0U, 0U, 0U}));  // A1, one body expansion
}

TEST(DepExtractor, LetBoundLambdaBodyIsExpandedOnCall) {
  Workbook wb = Workbook::create();
  Arena arena;
  const parser::AstNode* root = ParseFormula("LET(f,LAMBDA(x,A1+x),f(1))", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  ASSERT_EQ(deps.cell_deps.size(), 1U);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{0U, 0U, 0U}));  // A1 in invoked body
}

TEST(DepExtractor, NamedLambdaExpansionRestoresCallerLetBindings) {
  Workbook wb = Workbook::create();
  wb.set_defined_names({DefinedName{"Named", "LAMBDA(x,x+A1)", -1, false, ""}});
  Arena arena;
  const parser::AstNode* root = ParseFormula("LET(f,LAMBDA(x,B1+x),Named(1)+f(2))", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  std::vector<CellNodeId> expected = {
      CellNodeId{0U, 0U, 0U},  // Named body A1
      CellNodeId{0U, 0U, 1U},  // caller LET lambda body B1
  };
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, LambdaNamesShadowBuiltinVolatility) {
  Workbook wb = Workbook::create();
  wb.set_defined_names({DefinedName{"NOW", "LAMBDA(x,x+1)", -1, false, ""}});
  Arena arena;
  const parser::AstNode* root = ParseFormula("LET(RAND,LAMBDA(x,x),NOW(1)+RAND(2))", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
}

TEST(DepExtractor, NameRefCaseInsensitive) {
  Workbook wb = Workbook::create();
  // Defined name authored as `Foo`; the formula references `foo` (lowercase).
  // Excel resolves names case-insensitively, so the walker must match.
  std::vector<DefinedName> names;
  names.push_back(DefinedName{"Foo", "=A1", -1, false, ""});
  wb.set_defined_names(std::move(names));

  Arena arena;
  const parser::AstNode* root = ParseFormula("foo+1", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  ASSERT_EQ(deps.cell_deps.size(), 1u);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{0U, 0U, 0U}));
}

}  // namespace
}  // namespace formulon::eval
