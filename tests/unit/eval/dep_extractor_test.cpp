// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for the AST -> dep-graph adapter. The walker should:
//   * Resolve plain `Ref` nodes to single CellNodeIds on the bound sheet.
//   * Flatten Ref:Ref RangeOps into per-cell deps.
//   * Promote whole-column / whole-row references to volatile status without
//     enumerating cells.
//   * Detect Excel volatile functions inside any Call node.
//   * Resolve sheet-qualified refs against the workbook.
//   * Skip ExternalRef / NameRef / StructuredRef / Lambda body silently.
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
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
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

TEST(DepExtractor, WholeColumnRefIsVolatileWithNoCells) {
  Workbook wb = Workbook::create();
  Arena arena;
  // SUM(A:A) — whole-column. We do not enumerate the 1M cells; instead the
  // formula is promoted to volatile so dependents always recompute.
  const parser::AstNode* root = ParseFormula("SUM(A:A)", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_TRUE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, WholeRowRefIsVolatileWithNoCells) {
  Workbook wb = Workbook::create();
  Arena arena;
  const parser::AstNode* root = ParseFormula("SUM(1:1)", arena);
  ASSERT_NE(root, nullptr);
  ExtractedDeps deps = extract_deps(*root, 0U, wb);
  EXPECT_TRUE(deps.is_volatile);
  EXPECT_TRUE(deps.cell_deps.empty());
}

TEST(DepExtractor, NameRefIsSkipped) {
  Workbook wb = Workbook::create();
  Arena arena;
  // Bare identifier that is not a function call parses as NameRef. Defined
  // names are out of scope at this stage; the walker skips them.
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
