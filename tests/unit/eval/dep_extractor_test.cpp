//
// Unit tests for the AST -> dep-graph adapter. The walker should:
//   * Resolve plain `Ref` nodes to single CellNodeIds on the bound sheet.
//   * Flatten Ref:Ref RangeOps into per-cell deps.
//   * Promote whole-column / whole-row references to volatile status without
//     enumerating cells.
//   * Detect Excel volatile functions inside any Call node.
//   * Resolve sheet-qualified refs against the workbook.
//   * Resolve `NameRef` against the workbook's defined-name list (sheet-
//     scoped beats workbook-scoped, case-insensitive, cycles broken
//     silently, parser failures skipped).
//   * Resolve `StructuredRef` against the workbook's table metadata,
//     pinning the resulting rectangle on the table's owning sheet at
//     extract time. Implicit-intersection (`Table[@Col]`), unknown
//     tables, and unknown columns silently skip.
//   * Capture ExternalRef book ids in `external_book_ids` (deduplicated,
//     forward-compatible with future cross-workbook recalc wiring).
//   * Skip Lambda body silently.
//   * Descend into LET binding initialisers and the LET body, so refs and
//     volatile calls inside either surface as deps.

#include "eval/dep_extractor.h"

#include <algorithm>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/dep_graph.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "io/tables_reader.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/resource_budget.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Parses `source` into the supplied arena and returns the AST root. Asserts
// that parsing produced a usable tree (the walker assumes a non-null root).
const parser::AstNode* ParseFormula(std::string_view source, Arena& arena) {
  parser::Parser parser(source, arena);
  parser::AstNode* root = parser.parse();
  EXPECT_NE(root, nullptr);
  return root;
}

// Sorts a vector of CellNodeIds by (sheet, row, col) for stable comparison.
std::vector<CellNodeId> Sorted(std::vector<CellNodeId> v) {
  std::sort(v.begin(), v.end(), [](CellNodeId a, CellNodeId b) {
    if (a.sheet_id != b.sheet_id)
      return a.sheet_id < b.sheet_id;
    if (a.row != b.row)
      return a.row < b.row;
    return a.col < b.col;
  });
  return v;
}

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

TEST(DepExtractor, OffsetIsVolatileAndStillVisitsArgRefs) {
  Workbook wb = Workbook::create();
  Arena arena;
  // OFFSET(A1, 1, 1) — A1 is a literal Ref argument. Even though OFFSET
  // produces a dynamic ref at evaluation time, the literal anchor is a
  // direct dep at the AST level.
  const parser::AstNode* root = ParseFormula("OFFSET(A1,1,1)", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_TRUE(deps.is_volatile);
  // A1 is registered as a static dep of the OFFSET call.
  EXPECT_EQ(deps.cell_deps.size(), 1u);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{0U, 0U, 0U}));
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
  wb.set_defined_names({io::DefinedName{"Cols", "A:C", -1, false, ""},
                        io::DefinedName{"Apply", "LAMBDA(x,SUM(Cols)+x)", -1, false, ""}});

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

// --- Defined-name resolution ------------------------------------------------

TEST(DepExtractor, NameRefWorkbookScopedSingleCell) {
  Workbook wb = Workbook::create();
  // MyName -> =A1, workbook-scoped (local_sheet_id = -1).
  std::vector<io::DefinedName> names;
  names.push_back(io::DefinedName{"MyName", "=A1", -1, false, ""});
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
  std::vector<io::DefinedName> names;
  names.push_back(io::DefinedName{"MyRange", "=A1:B2", -1, false, ""});
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
  std::vector<io::DefinedName> names;
  names.push_back(io::DefinedName{"Name1", "=Name2", -1, false, ""});
  names.push_back(io::DefinedName{"Name2", "=A1", -1, false, ""});
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
  std::vector<io::DefinedName> names;
  names.push_back(io::DefinedName{"Foo", "=A1", -1, false, ""});
  names.push_back(io::DefinedName{"Foo", "=B1", 0, false, ""});
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
  std::vector<io::DefinedName> names;
  names.push_back(io::DefinedName{"Loop", "=Loop+1", -1, false, ""});
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
  std::vector<io::DefinedName> names;
  names.push_back(io::DefinedName{"RandName", "=RAND()", -1, false, ""});
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
  wb.set_defined_names({io::DefinedName{"Named", "LAMBDA(x,x+A1+RAND())", -1, false, ""}});
  Arena arena;
  const parser::AstNode* root = ParseFormula("Named(5)", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_TRUE(deps.is_volatile);
  ASSERT_EQ(deps.cell_deps.size(), 1U);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{0U, 0U, 0U}));  // A1
}

TEST(DepExtractor, ThreeDSpanMetadataIsNormalizedAndDeduplicatedAcrossExpansion) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "First")));
  wb.add_sheet("Second");
  wb.add_sheet("Third");
  wb.set_defined_names({
      io::DefinedName{"Inner", "First:Second!A1:A3", -1, false, ""},
      io::DefinedName{"Outer", "Inner", -1, false, ""},
      io::DefinedName{"F", "LAMBDA(x,SUM(Outer)+x)", -1, false, ""},
  });

  Arena arena;
  const parser::AstNode* root = ParseFormula("SUM(First:Second!A1:A3)+SUM(Second:First!A1:A3)+SUM(Outer)+F(0)", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);

  ASSERT_EQ(deps.three_d_spans.size(), 1U);
  EXPECT_EQ(deps.three_d_spans[0], (ThreeDSheetSpanDependency{0U, 1U}));
}

TEST(DepExtractor, ThreeDSpanMetadataIncludesWholeAxisReferences) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "First")));
  wb.add_sheet("Second");

  Arena arena;
  const parser::AstNode* root = ParseFormula("SUM(First:Second!A:A)+SUM(First:Second!1:1)", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);

  ASSERT_EQ(deps.three_d_spans.size(), 1U);
  EXPECT_EQ(deps.three_d_spans[0], (ThreeDSheetSpanDependency{0U, 1U}));
}

TEST(DepExtractor, NamedLambdaParameterShadowsDefinedName) {
  Workbook wb = Workbook::create();
  wb.set_defined_names({
      io::DefinedName{"x", "B1", -1, false, ""},
      io::DefinedName{"F", "LAMBDA(x,x+A1)", -1, false, ""},
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
      io::DefinedName{"x", "B1", -1, false, ""},
      io::DefinedName{"F", "LAMBDA(y,y+x)", -1, false, ""},
  });
  Arena arena;
  const parser::AstNode* root = ParseFormula("LET(x,A1,F(1))", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  std::vector<CellNodeId> expected = {
      CellNodeId{0U, 0U, 0U},  // LET initializer A1
      CellNodeId{0U, 0U, 1U},  // defined x -> B1 inside F
  };
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, NamedLambdaRecursiveBodyIsFinite) {
  Workbook wb = Workbook::create();
  wb.set_defined_names({io::DefinedName{"Fact", "LAMBDA(n,IF(n<=1,1,n*Fact(n-1)+A1))", -1, false, ""}});
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
  wb.set_defined_names({io::DefinedName{"Named", "LAMBDA(x,x+A1)", -1, false, ""}});
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
  wb.set_defined_names({io::DefinedName{"NOW", "LAMBDA(x,x+1)", -1, false, ""}});
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
  std::vector<io::DefinedName> names;
  names.push_back(io::DefinedName{"Foo", "=A1", -1, false, ""});
  wb.set_defined_names(std::move(names));

  Arena arena;
  const parser::AstNode* root = ParseFormula("foo+1", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  ASSERT_EQ(deps.cell_deps.size(), 1u);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{0U, 0U, 0U}));
}

// --- Structured (table) references ------------------------------------------

// Builds a `TableMetadata` with `display_name == name` and the given column
// names. `ref` is the raw A1 footprint including header and totals rows.
io::TableMetadata MakeTable(std::string name, std::string ref, std::size_t sheet_index, bool header_row,
                            bool totals_row, std::vector<std::string> column_names) {
  io::TableMetadata t;
  t.id = 1;
  t.name = name;
  t.display_name = std::move(name);
  t.ref = std::move(ref);
  t.sheet_index = sheet_index;
  t.header_row = header_row;
  t.totals_row = totals_row;
  t.columns.reserve(column_names.size());
  std::uint32_t next_id = 1;
  for (auto& cname : column_names) {
    io::TableColumn col;
    col.id = next_id++;
    col.name = std::move(cname);
    t.columns.push_back(std::move(col));
  }
  return t;
}

TEST(DepExtractor, StructuredRefDefaultModifierFlattensDataColumn) {
  Workbook wb = Workbook::create();
  // Sales table at A1:C10 with header, no totals, columns Region/Amount/Date.
  // SUM(Sales[Amount]) defaults to the kData area on column 1 (Amount), so
  // deps must be B2..B10.
  std::vector<io::TableMetadata> tables;
  tables.push_back(MakeTable("Sales", "A1:C10", /*sheet_index=*/0, /*header_row=*/true, /*totals_row=*/false,
                             {"Region", "Amount", "Date"}));
  wb.set_tables(std::move(tables));

  Arena arena;
  const parser::AstNode* root = ParseFormula("SUM(Sales[Amount])", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, /*current_sheet_id=*/0U, wb);
  EXPECT_FALSE(deps.is_volatile);

  std::vector<CellNodeId> expected;
  for (std::uint32_t r = 1; r <= 9; ++r) {
    expected.push_back(CellNodeId{0U, r, 1U});  // B2..B10
  }
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, StructuredRefAboveLimitUsesCompactRangeDependency) {
  Workbook wb = Workbook::create();
  // A table column is as tall as the table, so the structured-ref path needs
  // the same graph-footprint ceiling the RangeOp path has: the data area of
  // this column is 4,000 cells.
  std::vector<io::TableMetadata> tables;
  tables.push_back(MakeTable("Sales", "A1:C4001", /*sheet_index=*/0, /*header_row=*/true, /*totals_row=*/false,
                             {"Region", "Amount", "Date"}));
  wb.set_tables(std::move(tables));

  Arena arena;
  const parser::AstNode* root = ParseFormula("SUM(Sales[Amount])", arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, /*current_sheet_id=*/0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
  ASSERT_EQ(deps.range_deps.size(), 1U);
  EXPECT_EQ(deps.range_deps[0].sheet_id, 0U);
  EXPECT_EQ(deps.range_deps[0].row_first, 1U);
  EXPECT_EQ(deps.range_deps[0].row_last, 4000U);
  EXPECT_EQ(deps.range_deps[0].col_first, 1U);
  EXPECT_EQ(deps.range_deps[0].col_last, 1U);
}

TEST(DepExtractor, StructuredRefDataExcludesTotalsRow) {
  Workbook wb = Workbook::create();
  // Same table, now with a totals row. kData must skip both header (row 0)
  // and totals (row 9), so deps are B2..B9.
  std::vector<io::TableMetadata> tables;
  tables.push_back(MakeTable("Sales", "A1:C10", /*sheet_index=*/0, /*header_row=*/true, /*totals_row=*/true,
                             {"Region", "Amount", "Date"}));
  wb.set_tables(std::move(tables));

  Arena arena;
  const parser::AstNode* root = ParseFormula("SUM(Sales[Amount])", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);

  std::vector<CellNodeId> expected;
  for (std::uint32_t r = 1; r <= 8; ++r) {
    expected.push_back(CellNodeId{0U, r, 1U});  // B2..B9
  }
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, StructuredRefHeadersOnly) {
  Workbook wb = Workbook::create();
  // Sales[#Headers] -> A1..C1.
  std::vector<io::TableMetadata> tables;
  tables.push_back(MakeTable("Sales", "A1:C10", /*sheet_index=*/0, /*header_row=*/true, /*totals_row=*/false,
                             {"Region", "Amount", "Date"}));
  wb.set_tables(std::move(tables));

  Arena arena;
  const parser::AstNode* root = ParseFormula("Sales[#Headers]", arena);
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

TEST(DepExtractor, StructuredRefTotalsOnly) {
  Workbook wb = Workbook::create();
  // Table with totals. Sales[#Totals] -> last row of `ref`, i.e. A10..C10.
  std::vector<io::TableMetadata> tables;
  tables.push_back(MakeTable("Sales", "A1:C10", /*sheet_index=*/0, /*header_row=*/true, /*totals_row=*/true,
                             {"Region", "Amount", "Date"}));
  wb.set_tables(std::move(tables));

  Arena arena;
  const parser::AstNode* root = ParseFormula("Sales[#Totals]", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);

  std::vector<CellNodeId> expected = {
      CellNodeId{0U, 9U, 0U},  // A10
      CellNodeId{0U, 9U, 1U},  // B10
      CellNodeId{0U, 9U, 2U},  // C10
  };
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, StructuredRefAllCoversFullRectangle) {
  Workbook wb = Workbook::create();
  // Sales[#All] -> every row in the ref rect (A1..C10).
  std::vector<io::TableMetadata> tables;
  tables.push_back(MakeTable("Sales", "A1:C10", /*sheet_index=*/0, /*header_row=*/true, /*totals_row=*/false,
                             {"Region", "Amount", "Date"}));
  wb.set_tables(std::move(tables));

  Arena arena;
  const parser::AstNode* root = ParseFormula("Sales[#All]", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);

  std::vector<CellNodeId> expected;
  for (std::uint32_t r = 0; r < 10; ++r) {
    for (std::uint32_t c = 0; c < 3; ++c) {
      expected.push_back(CellNodeId{0U, r, c});
    }
  }
  EXPECT_EQ(deps.cell_deps.size(), expected.size());
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, StructuredRefAllAreaWithSpecificColumn) {
  Workbook wb = Workbook::create();
  // Sales[[#All],[Amount]] -> column 1 (Amount) across every row of the ref
  // rectangle, including header. Excel emits the bracket payload as
  // `[#All],[Amount]` so the parser stores that verbatim.
  std::vector<io::TableMetadata> tables;
  tables.push_back(MakeTable("Sales", "A1:C10", /*sheet_index=*/0, /*header_row=*/true, /*totals_row=*/false,
                             {"Region", "Amount", "Date"}));
  wb.set_tables(std::move(tables));

  Arena arena;
  const parser::AstNode* root = ParseFormula("Sales[[#All],[Amount]]", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);

  std::vector<CellNodeId> expected;
  for (std::uint32_t r = 0; r < 10; ++r) {
    expected.push_back(CellNodeId{0U, r, 1U});  // B1..B10
  }
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
}

TEST(DepExtractor, StructuredRefImplicitIntersectionSilentSkip) {
  Workbook wb = Workbook::create();
  // Sales[@Amount] resolves at evaluator time to a single cell on the
  // formula's row. The dep extractor cannot know the formula's row, so
  // it silently skips — the evaluator surfaces the actual dep when the
  // implicit intersection resolves at eval time.
  std::vector<io::TableMetadata> tables;
  tables.push_back(MakeTable("Sales", "A1:C10", /*sheet_index=*/0, /*header_row=*/true, /*totals_row=*/false,
                             {"Region", "Amount", "Date"}));
  wb.set_tables(std::move(tables));

  Arena arena;
  const parser::AstNode* root = ParseFormula("Sales[@Amount]", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, StructuredRefUnknownTableSilentSkip) {
  Workbook wb = Workbook::create();
  // No tables registered: NoSuchTable[Col] cannot resolve and produces
  // no deps. Same silent-skip policy as unknown sheet qualifiers.
  Arena arena;
  const parser::AstNode* root = ParseFormula("NoSuchTable[Col]", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, StructuredRefUnknownColumnSilentSkip) {
  Workbook wb = Workbook::create();
  // Table exists, column does not: silent skip, no deps.
  std::vector<io::TableMetadata> tables;
  tables.push_back(MakeTable("Sales", "A1:C10", /*sheet_index=*/0, /*header_row=*/true, /*totals_row=*/false,
                             {"Region", "Amount", "Date"}));
  wb.set_tables(std::move(tables));

  Arena arena;
  const parser::AstNode* root = ParseFormula("Sales[NotAColumn]", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, StructuredRefCrossSheetTableLandsOnTableSheet) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");  // index 1
  wb.add_sheet("Sheet3");  // index 2
  // Table sits on sheet 2; formula is being analysed for a cell on sheet 0.
  // Deps must land on sheet 2.
  std::vector<io::TableMetadata> tables;
  tables.push_back(MakeTable("Sales", "A1:C10", /*sheet_index=*/2, /*header_row=*/true, /*totals_row=*/false,
                             {"Region", "Amount", "Date"}));
  wb.set_tables(std::move(tables));

  Arena arena;
  const parser::AstNode* root = ParseFormula("SUM(Sales[Amount])", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, /*current_sheet_id=*/0U, wb);
  EXPECT_FALSE(deps.is_volatile);

  std::vector<CellNodeId> expected;
  for (std::uint32_t r = 1; r <= 9; ++r) {
    expected.push_back(CellNodeId{2U, r, 1U});  // Sheet3!B2..B10
  }
  EXPECT_EQ(Sorted(deps.cell_deps), Sorted(expected));
  // Defensive: every dep must carry sheet_id == 2.
  for (const CellNodeId& id : deps.cell_deps) {
    EXPECT_EQ(id.sheet_id, 2U);
  }
}

TEST(DepExtractor, StructuredRefHeadersOnHeaderlessTableSilentSkip) {
  Workbook wb = Workbook::create();
  // header_row=false: the table has no header band. Sales[#Headers] is
  // unresolvable; the resolver returns ErrorCode::Ref and we silent-skip.
  std::vector<io::TableMetadata> tables;
  tables.push_back(MakeTable("Sales", "A1:C10", /*sheet_index=*/0, /*header_row=*/false, /*totals_row=*/false,
                             {"Region", "Amount", "Date"}));
  wb.set_tables(std::move(tables));

  Arena arena;
  const parser::AstNode* root = ParseFormula("Sales[#Headers]", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

// ---------------------------------------------------------------------------
// ExternalRef opaque-sentinel tracking
// ---------------------------------------------------------------------------

TEST(DepExtractor, ExternalRefRecordsBookId) {
  Workbook wb = Workbook::create();
  Arena arena;
  parser::Reference cell{};
  cell.row = 0U;
  cell.col = 0U;
  parser::AstNode* root = parser::make_external_ref(arena, /*book_id=*/3U, "Sheet1", cell);
  ASSERT_NE(root, nullptr);

  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
  ASSERT_EQ(deps.external_book_ids.size(), 1u);
  EXPECT_EQ(deps.external_book_ids[0], 3u);
}

TEST(DepExtractor, ExternalRefDeduplicatesRepeatedBook) {
  Workbook wb = Workbook::create();
  Arena arena;
  parser::Reference cell_a{};
  cell_a.row = 0U;
  cell_a.col = 0U;
  parser::Reference cell_b{};
  cell_b.row = 1U;
  cell_b.col = 1U;
  parser::AstNode* lhs = parser::make_external_ref(arena, /*book_id=*/2U, "Sheet1", cell_a);
  parser::AstNode* rhs = parser::make_external_ref(arena, /*book_id=*/2U, "Sheet1", cell_b);
  ASSERT_NE(lhs, nullptr);
  ASSERT_NE(rhs, nullptr);
  parser::AstNode* root = parser::make_binary_op(arena, parser::BinOp::Add, lhs, rhs);
  ASSERT_NE(root, nullptr);

  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_TRUE(deps.cell_deps.empty());
  ASSERT_EQ(deps.external_book_ids.size(), 1u);
  EXPECT_EQ(deps.external_book_ids[0], 2u);
}

TEST(DepExtractor, ExternalRefMultipleBooksPreservedInOrder) {
  Workbook wb = Workbook::create();
  Arena arena;
  parser::Reference cell{};
  cell.row = 0U;
  cell.col = 0U;
  parser::AstNode* a = parser::make_external_ref(arena, /*book_id=*/5U, "Sheet1", cell);
  parser::AstNode* b = parser::make_external_ref(arena, /*book_id=*/1U, "Sheet1", cell);
  parser::AstNode* root = parser::make_binary_op(arena, parser::BinOp::Add, a, b);
  ASSERT_NE(root, nullptr);

  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  ASSERT_EQ(deps.external_book_ids.size(), 2u);
  // First-encounter order; the walker sees `a` (5) before `b` (1).
  EXPECT_EQ(deps.external_book_ids[0], 5u);
  EXPECT_EQ(deps.external_book_ids[1], 1u);
}

TEST(DepExtractor, ExternalRefDoesNotEnumerateCells) {
  // ExternalRef must contribute zero CellNodeId entries — its sheet lives
  // outside the workbook's sheet table, so a CellNodeId would index the
  // wrong sheet. Combining with a local Ref lets us confirm only the local
  // cell surfaces in cell_deps.
  Workbook wb = Workbook::create();
  Arena arena;
  parser::Reference ext_cell{};
  ext_cell.row = 5U;
  ext_cell.col = 7U;
  parser::AstNode* ext = parser::make_external_ref(arena, /*book_id=*/1U, "Sheet1", ext_cell);

  parser::Reference b2{};
  b2.row = 1U;
  b2.col = 1U;
  parser::AstNode* b2_node = parser::make_ref(arena, b2);
  parser::AstNode* root = parser::make_binary_op(arena, parser::BinOp::Add, ext, b2_node);
  ASSERT_NE(root, nullptr);

  ExtractedDeps deps = extract_deps(*root, /*current_sheet_id=*/0U, wb);
  ASSERT_EQ(deps.cell_deps.size(), 1u);
  EXPECT_EQ(deps.cell_deps[0], (CellNodeId{0U, 1U, 1U}));
  ASSERT_EQ(deps.external_book_ids.size(), 1u);
  EXPECT_EQ(deps.external_book_ids[0], 1u);
}

TEST(DepExtractor, ExternalRefIsNotMarkedVolatile) {
  // Capturing the book id is enough; the formula is not volatile in the
  // RAND/NOW sense.
  Workbook wb = Workbook::create();
  Arena arena;
  parser::Reference cell{};
  parser::AstNode* root = parser::make_external_ref(arena, /*book_id=*/9U, "Sheet1", cell);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
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

TEST(DepExtractor, ThreeDFullColumnUsesCompactRangeDependencies) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.add_sheet("Sheet3");
  Arena arena;
  const parser::AstNode* root = ParseFormula("SUM(Sheet1:Sheet3!A:A)", arena);
  ASSERT_NE(root, nullptr);

  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
  ASSERT_EQ(deps.range_deps.size(), 3U);
  for (std::size_t i = 0; i < deps.range_deps.size(); ++i) {
    EXPECT_EQ(deps.range_deps[i].sheet_id, i);
    EXPECT_EQ(deps.range_deps[i].row_last, Sheet::kMaxRows - 1U);
    EXPECT_EQ(deps.range_deps[i].col_first, 0U);
    EXPECT_EQ(deps.range_deps[i].col_last, 0U);
  }
}

TEST(DepExtractor, ThreeDRangeAboveLimitUsesCompactRangeDependencyPerSheet) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.add_sheet("Sheet3");
  Arena arena;
  // The shared rectangle is read once per sheet in the span, so the ceiling
  // is applied before the span multiplies the cost: 3 x 2,000 cells would be
  // 6,000 permanent edges from one formula.
  const parser::AstNode* root = ParseFormula("SUM(Sheet1:Sheet3!A1:B1000)", arena);
  ASSERT_NE(root, nullptr);

  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_FALSE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
  ASSERT_EQ(deps.range_deps.size(), 3U);
  for (std::size_t sheet = 0; sheet < deps.range_deps.size(); ++sheet) {
    EXPECT_EQ(deps.range_deps[sheet].sheet_id, sheet);
    EXPECT_EQ(deps.range_deps[sheet].row_first, 0U);
    EXPECT_EQ(deps.range_deps[sheet].row_last, 999U);
    EXPECT_EQ(deps.range_deps[sheet].col_first, 0U);
    EXPECT_EQ(deps.range_deps[sheet].col_last, 1U);
  }
}

TEST(DepExtractor, ThreeDWholeAxisSpansPreserveBothAxes) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.add_sheet("Sheet3");
  Arena arena;
  const parser::AstNode* root = ParseFormula("SUM(Sheet1:Sheet3!A:C)+SUM(Sheet1:Sheet3!1:3)", arena);
  ASSERT_NE(root, nullptr);

  const ExtractedDeps deps = extract_deps(*root, 0U, wb);
  ASSERT_EQ(deps.range_deps.size(), 6U);
  for (std::size_t sheet = 0; sheet < 3U; ++sheet) {
    const CellRangeDependency& columns = deps.range_deps[sheet];
    EXPECT_EQ(columns.sheet_id, sheet);
    EXPECT_EQ(columns.row_first, 0U);
    EXPECT_EQ(columns.row_last, Sheet::kMaxRows - 1U);
    EXPECT_EQ(columns.col_first, 0U);
    EXPECT_EQ(columns.col_last, 2U);

    const CellRangeDependency& rows = deps.range_deps[3U + sheet];
    EXPECT_EQ(rows.sheet_id, sheet);
    EXPECT_EQ(rows.row_first, 0U);
    EXPECT_EQ(rows.row_last, 2U);
    EXPECT_EQ(rows.col_first, 0U);
    EXPECT_EQ(rows.col_last, Sheet::kMaxCols - 1U);
  }
}

}  // namespace
}  // namespace formulon::eval
