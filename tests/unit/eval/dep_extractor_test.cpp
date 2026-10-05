// Dependency extractor tests grouped by dependency source.

#include "eval/dep_extractor.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "defined_name.h"
#include "dep_extractor_test_helpers.h"
#include "table.h"
#include "utils/resource_budget.h"
#include "workbook.h"

namespace formulon::eval {
namespace {
using namespace dep_extractor_test_helpers;

TEST(DepExtractor, BinaryAddTwoCellRefs) {
  Workbook wb = Workbook::create();
  Arena arena;
  const parser::AstNode* root = ParseFormula("A1+B2", arena);
  ASSERT_NE(root, nullptr);

  ExtractedDeps deps = extract_deps(*root, /*current_sheet_id=*/0U, wb);

  EXPECT_FALSE(deps.is_volatile);
  EXPECT_EQ(deps.cell_deps.size(), 2u);
  std::vector<CellNodeId> expected = {
      CellNodeId{0U, 0U, 0U},  // A1
      CellNodeId{0U, 1U, 1U},  // B2
  };
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, RangeOpFlattensCells) {
  Workbook wb = Workbook::create();
  Arena arena;
  const parser::AstNode* root = ParseFormula("SUM(A1:A3)", arena);
  ASSERT_NE(root, nullptr);

  ExtractedDeps deps = extract_deps(*root, /*current_sheet_id=*/0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_EQ(deps.cell_deps.size(), 3u);
  std::vector<CellNodeId> expected = {
      CellNodeId{0U, 0U, 0U},
      CellNodeId{0U, 1U, 0U},
      CellNodeId{0U, 2U, 0U},
  };
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, NowIsVolatile) {
  Workbook wb = Workbook::create();
  Arena arena;
  const parser::AstNode* root = ParseFormula("NOW()", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_TRUE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, LowerCaseNowIsVolatile) {
  // A hand-typed `=now()` keeps its lowercase lexeme; volatile detection
  // is case-insensitive, so the cell must still be flagged volatile (and
  // thus re-fire on every recalc).
  Workbook wb = Workbook::create();
  Arena arena;
  const parser::AstNode* root = ParseFormula("now()", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_TRUE(deps.is_volatile);
}

TEST(DepExtractor, MixedCaseOffsetIsVolatile) {
  Workbook wb = Workbook::create();
  Arena arena;
  const parser::AstNode* root = ParseFormula("Offset(A1,1,1)", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_TRUE(deps.is_volatile);
}

TEST(DepExtractor, RandIsVolatile) {
  Workbook wb = Workbook::create();
  Arena arena;
  const parser::AstNode* root = ParseFormula("RAND()", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_TRUE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, ValueVolatileCallsCarryNoDynamicReference) {
  // These read the clock, the RNG or workbook metadata; the cells they
  // read (none, or a reference the caller spelled out) are exactly what
  // `cell_deps` records.
  Workbook wb = Workbook::create();
  for (const char* formula :
       {"NOW()", "TODAY()", "RAND()", "RANDBETWEEN(1,9)", "RANDARRAY(2,2)", "INFO(\"numfile\")", "CELL(\"row\",A1)"}) {
    Arena arena;
    const parser::AstNode* root = ParseFormula(formula, arena);
    ASSERT_NE(root, nullptr) << formula;
    ExtractedDeps deps = extract_deps(*root, 0U, wb);
    EXPECT_TRUE(deps.is_volatile) << formula;
    EXPECT_FALSE(deps.has_dynamic_reference) << formula;
  }
}

TEST(DepExtractor, IndirectAndOffsetCarryADynamicReference) {
  Workbook wb = Workbook::create();
  for (const char* formula :
       {"INDIRECT(\"A1\")", "OFFSET(A1,1,1)", "indirect(B2&\"\")", "Offset(A1,1,1)", "SUM(RAND(),INDIRECT(\"A1\"))"}) {
    Arena arena;
    const parser::AstNode* root = ParseFormula(formula, arena);
    ASSERT_NE(root, nullptr) << formula;
    ExtractedDeps deps = extract_deps(*root, 0U, wb);
    EXPECT_TRUE(deps.is_volatile) << formula;
    EXPECT_TRUE(deps.has_dynamic_reference) << formula;
  }
}

TEST(DepExtractor, ShadowedIndirectCarriesNoDynamicReference) {
  // A LET-bound name that merely spells like `INDIRECT` is not the
  // built-in, so it must not force the cell onto the serial path.
  Workbook wb = Workbook::create();
  Arena arena;
  const parser::AstNode* root = ParseFormula("LET(INDIRECT,LAMBDA(x,x+1),INDIRECT(1))", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_FALSE(deps.has_dynamic_reference);
}

TEST(DepExtractor, MixedScalarAndRangeNonVolatile) {
  Workbook wb = Workbook::create();
  Arena arena;
  // A1 + SUM(B1:B3) -> 4 deps total: A1, B1, B2, B3.
  const parser::AstNode* root = ParseFormula("A1+SUM(B1:B3)", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  std::vector<CellNodeId> expected = {
      CellNodeId{0U, 0U, 0U},  // A1
      CellNodeId{0U, 0U, 1U},  // B1
      CellNodeId{0U, 1U, 1U},  // B2
      CellNodeId{0U, 2U, 1U},  // B3
  };
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, CrossSheetRefResolvesToTargetSheetId) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");  // index 1
  Arena arena;
  const parser::AstNode* root = ParseFormula("Sheet2!A1", arena);
  ASSERT_NE(root, nullptr);
  // The formula lives on Sheet1 (index 0) but reads from Sheet2 (index 1).
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  ASSERT_EQ(deps.cell_deps.size(), 1u);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{1U, 0U, 0U}));
}

TEST(DepExtractor, UnknownSheetRefIsSkipped) {
  Workbook wb = Workbook::create();
  Arena arena;
  // GhostSheet does not exist; the walker should drop the ref silently
  // rather than emitting a bogus CellNodeId on sheet 0.
  const parser::AstNode* root = ParseFormula("GhostSheet!A1", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, WholeColumnRefUsesCompactRangeDependency) {
  Workbook wb = Workbook::create();
  Arena arena;
  // SUM(A:A) — whole-column. We do not enumerate its 1M cells or make a
  // non-volatile formula re-execute on every recalc pass.
  const parser::AstNode* root = ParseFormula("SUM(A:A)", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
  ASSERT_EQ(deps.range_deps.size(), 1U);
  EXPECT_EQ(deps.range_deps[0].sheet_id, 0U);
  EXPECT_EQ(deps.range_deps[0].row_first, 0U);
  EXPECT_EQ(deps.range_deps[0].row_last, Sheet::kMaxRows - 1U);
  EXPECT_EQ(deps.range_deps[0].col_first, 0U);
  EXPECT_EQ(deps.range_deps[0].col_last, 0U);
}

TEST(DepExtractor, WholeRowRefUsesCompactRangeDependency) {
  Workbook wb = Workbook::create();
  Arena arena;
  const parser::AstNode* root = ParseFormula("SUM(1:1)", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
  ASSERT_EQ(deps.range_deps.size(), 1U);
  EXPECT_EQ(deps.range_deps[0].row_first, 0U);
  EXPECT_EQ(deps.range_deps[0].row_last, 0U);
  EXPECT_EQ(deps.range_deps[0].col_first, 0U);
  EXPECT_EQ(deps.range_deps[0].col_last, Sheet::kMaxCols - 1U);
}

TEST(DepExtractor, WholeAxisSpanNormalizesEveryAstForm) {
  Workbook wb = Workbook::create();
  Arena arena;

  const parser::AstNode* columns = ParseFormula("SUM(A:C)", arena);
  ASSERT_NE(columns, nullptr);
  const ExtractedDeps column_deps = extract_deps(*columns, 0U, wb);
  ASSERT_EQ(column_deps.range_deps.size(), 1U);
  EXPECT_EQ(column_deps.range_deps[0].row_first, 0U);
  EXPECT_EQ(column_deps.range_deps[0].row_last, Sheet::kMaxRows - 1U);
  EXPECT_EQ(column_deps.range_deps[0].col_first, 0U);
  EXPECT_EQ(column_deps.range_deps[0].col_last, 2U);

  arena.reset();
  const parser::AstNode* rows = ParseFormula("SUM(1:3)", arena);
  ASSERT_NE(rows, nullptr);
  const ExtractedDeps row_deps = extract_deps(*rows, 0U, wb);
  ASSERT_EQ(row_deps.range_deps.size(), 1U);
  EXPECT_EQ(row_deps.range_deps[0].row_first, 0U);
  EXPECT_EQ(row_deps.range_deps[0].row_last, 2U);
  EXPECT_EQ(row_deps.range_deps[0].col_first, 0U);
  EXPECT_EQ(row_deps.range_deps[0].col_last, Sheet::kMaxCols - 1U);
}

TEST(DepExtractor, WholeAxisSpanNormalizationReachesNamesAndLambdas) {
  Workbook wb = Workbook::create();
  wb.set_defined_names(
      {DefinedName{"Cols", "A:C", -1, false, ""}, DefinedName{"Apply", "LAMBDA(x,SUM(Cols)+x)", -1, false, ""}});

  Arena arena;
  const parser::AstNode* root = ParseFormula("Apply(1)", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  ASSERT_EQ(deps.range_deps.size(), 1U);
  EXPECT_EQ(deps.range_deps[0].row_first, 0U);
  EXPECT_EQ(deps.range_deps[0].row_last, Sheet::kMaxRows - 1U);
  EXPECT_EQ(deps.range_deps[0].col_first, 0U);
  EXPECT_EQ(deps.range_deps[0].col_last, 2U);
}

TEST(DepExtractor, BoundedRectAtLimitStillFlattens) {
  Workbook wb = Workbook::create();
  Arena arena;
  // Exactly `kMaxMaterializedDependencyCells` cells: the ceiling is
  // inclusive, so this rectangle keeps its per-cell edges and the exact
  // ordering guarantees they carry.
  const parser::AstNode* root = ParseFormula("SUM(A1:A1024)", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.range_deps.empty());
  EXPECT_EQ(deps.cell_deps.size(), kMaxMaterializedDependencyCells);
}

TEST(DepExtractor, BoundedRectAboveLimitUsesCompactRangeDependency) {
  Workbook wb = Workbook::create();
  Arena arena;
  // One cell past the ceiling. Flattening a lookup table costs one permanent
  // graph edge per cell in three indexes, so the rectangle is retained whole
  // instead.
  const parser::AstNode* root = ParseFormula("SUM(A1:A1025)", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
  ASSERT_EQ(deps.range_deps.size(), 1U);
  EXPECT_EQ(deps.range_deps[0].sheet_id, 0U);
  EXPECT_EQ(deps.range_deps[0].row_first, 0U);
  EXPECT_EQ(deps.range_deps[0].row_last, 1024U);
  EXPECT_EQ(deps.range_deps[0].col_first, 0U);
  EXPECT_EQ(deps.range_deps[0].col_last, 0U);
}

TEST(DepExtractor, OversizedRectUsesCompactRangeDependency) {
  Workbook wb = Workbook::create();
  Arena arena;
  // This is a valid bounded rectangle, but expands to the entire grid.
  // Registering all of its direct dependencies would allocate ~17 billion
  // graph nodes while merely loading a workbook. It is not volatile either:
  // volatility would re-execute the formula on every pass whether or not
  // anything it reads changed, and would still leave a write inside the
  // rectangle unable to dirty anything on its own. The compact rectangle
  // keeps the dependency real at a fixed cost.
  const parser::AstNode* root = ParseFormula("SUM(A1:XFD1048576)", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
  ASSERT_EQ(deps.range_deps.size(), 1U);
  EXPECT_EQ(deps.range_deps[0].row_first, 0U);
  EXPECT_EQ(deps.range_deps[0].row_last, Sheet::kMaxRows - 1U);
  EXPECT_EQ(deps.range_deps[0].col_first, 0U);
  EXPECT_EQ(deps.range_deps[0].col_last, Sheet::kMaxCols - 1U);
}

TEST(DepExtractor, UnresolvedNameRefIsSkipped) {
  Workbook wb = Workbook::create();
  Arena arena;
  // Bare identifier that is not a function call parses as NameRef. With no
  // matching defined name in the workbook the walker silently skips it,
  // mirroring the policy used for unknown sheet qualifiers.
  const parser::AstNode* root = ParseFormula("MyName", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, ArrayLiteralIsTraversedButContainsOnlyLiterals) {
  // Excel's inline array literals (`{1,2;3,4}`) only allow scalar literals,
  // not cell refs — the parser rejects `{A1,B1}` as a syntax error. Confirm
  // that a literal-only array contributes no cell deps and is non-volatile.
  Workbook wb = Workbook::create();
  Arena arena;
  const parser::AstNode* root = ParseFormula("{1,2;3,4}", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, DuplicateRefDeduplicated) {
  Workbook wb = Workbook::create();
  Arena arena;
  // A1+A1+A1 — the walker should emit A1 only once.
  const parser::AstNode* root = ParseFormula("A1+A1+A1", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  ASSERT_EQ(deps.cell_deps.size(), 1u);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{0U, 0U, 0U}));
}

TEST(DepExtractor, NestedVolatileInsideArithmetic) {
  Workbook wb = Workbook::create();
  Arena arena;
  // 1 + RAND() — volatility should still bubble out from the inner call.
  const parser::AstNode* root = ParseFormula("1+RAND()", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_TRUE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, LambdaBodyNotDescendedInto) {
  Workbook wb = Workbook::create();
  Arena arena;
  // LAMBDA(x, x + A1) — A1 is a free reference inside the lambda body. The
  // walker intentionally does not descend into lambda bodies (binding-time
  // capture analysis is out of scope), so cell_deps is empty.
  const parser::AstNode* root = ParseFormula("LAMBDA(x,x+A1)", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, LetBindingInitialiserAndBodyBothEmitCellDeps) {
  Workbook wb = Workbook::create();
  Arena arena;
  // =LET(x, A1, x + B1) — the binding initialiser reads A1; the body
  // references the bound name `x` (which contributes nothing today since
  // NameRef is a no-op) plus a real cell `B1`. Both A1 and B1 must surface
  // as static deps so the recalc engine re-runs the LET formula when either
  // changes.
  const parser::AstNode* root = ParseFormula("LET(x,A1,x+B1)", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  std::vector<CellNodeId> expected = {
      CellNodeId{0U, 0U, 0U},  // A1
      CellNodeId{0U, 0U, 1U},  // B1
  };
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, LetBindingBoundNameDoesNotEmitDeps) {
  Workbook wb = Workbook::create();
  Arena arena;
  // =LET(x, A1, x) — the body is the bound name itself. The NameRef case is
  // a no-op pending defined-name support, so the body contributes nothing
  // and only the initialiser's A1 is recorded.
  const parser::AstNode* root = ParseFormula("LET(x,A1,x)", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  ASSERT_EQ(deps.cell_deps.size(), 1u);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{0U, 0U, 0U}));
}

TEST(DepExtractor, LetBodyVolatileCallIsDetected) {
  Workbook wb = Workbook::create();
  Arena arena;
  // =LET(x, 1, x + RAND()) — the volatile call lives in the body. With body
  // descent, RAND() must promote the formula to volatile.
  const parser::AstNode* root = ParseFormula("LET(x,1,x+RAND())", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_TRUE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, LetNestedLetBodiesAreDescended) {
  Workbook wb = Workbook::create();
  Arena arena;
  // =LET(a, A1, LET(b, B1, a + b + C1)) — the outer body is itself a LET
  // whose body references a real cell C1 and the bound names a, b. All
  // three cell deps must surface through both layers of body descent.
  const parser::AstNode* root = ParseFormula("LET(a,A1,LET(b,B1,a+b+C1))", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  std::vector<CellNodeId> expected = {
      CellNodeId{0U, 0U, 0U},  // A1
      CellNodeId{0U, 0U, 1U},  // B1
      CellNodeId{0U, 0U, 2U},  // C1
  };
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, RangeAcrossSheetQualifierOnLeft) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");  // index 1
  Arena arena;
  // Sheet2!A1:B2 — the parser keeps the qualifier on the LHS only; the RHS
  // inherits. All four cells should land on sheet 1.
  const parser::AstNode* root = ParseFormula("SUM(Sheet2!A1:B2)", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  std::vector<CellNodeId> expected = {
      CellNodeId{1U, 0U, 0U},
      CellNodeId{1U, 0U, 1U},  //
      CellNodeId{1U, 1U, 0U},
      CellNodeId{1U, 1U, 1U},
  };
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

}  // namespace
}  // namespace formulon::eval
