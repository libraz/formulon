// Dependency extractor tests grouped by dependency source.

#include <algorithm>
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

std::vector<CellNodeId> RectCells(std::uint16_t sheet, std::uint32_t r0, std::uint32_t r1, std::uint32_t c0,
                                  std::uint32_t c1) {
  std::vector<CellNodeId> out;
  for (std::uint32_t r = r0; r <= r1; ++r) {
    for (std::uint32_t c = c0; c <= c1; ++c) {
      out.push_back(CellNodeId{sheet, r, c});
    }
  }
  return out;
}

ExtractedDeps ExtractFrom(std::string_view source, const Workbook& wb) {
  Arena arena;
  const parser::AstNode* root = ParseFormula(source, arena);
  EXPECT_NE(root, nullptr);
  return root == nullptr ? ExtractedDeps{} : extract_deps(*root, 0U, wb);
}

bool HasRangeDep(const ExtractedDeps& deps, CellRangeDependency want) {
  return std::any_of(deps.range_deps.begin(), deps.range_deps.end(), [&](const CellRangeDependency& d) {
    return d.sheet_id == want.sheet_id && d.row_first == want.row_first && d.row_last == want.row_last &&
           d.col_first == want.col_first && d.col_last == want.col_last;
  });
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

TEST(DepExtractor, DynamicEndpointInsideSourceAddsNothing) {
  const Workbook wb = Workbook::create();
  const ExtractedDeps deps = ExtractFrom("SUM(A1:INDEX(A1:A10,3))", wb);
  EXPECT_EQ(Sorted(deps.cell_deps), RectCells(0U, 0U, 9U, 0U, 0U));
  EXPECT_TRUE(deps.range_deps.empty());
}

TEST(DepExtractor, DynamicEndpointRegistersBoundingBox) {
  const Workbook wb = Workbook::create();
  EXPECT_EQ(Sorted(ExtractFrom("SUM(A1:INDEX(C1:C10,3))", wb).cell_deps), RectCells(0U, 0U, 9U, 0U, 2U));
  EXPECT_EQ(Sorted(ExtractFrom("SUM(INDEX(A1:A10,2):INDEX(C1:C10,3))", wb).cell_deps), RectCells(0U, 0U, 9U, 0U, 2U));
  EXPECT_EQ(Sorted(ExtractFrom("SUM(A1:CHOOSE(2,B1,D5))", wb).cell_deps), RectCells(0U, 0U, 4U, 0U, 3U));
}

TEST(DepExtractor, ChainedRangeRegistersBoundingBox) {
  const Workbook wb = Workbook::create();
  EXPECT_EQ(Sorted(ExtractFrom("SUM(A1:B2:C3)", wb).cell_deps), RectCells(0U, 0U, 2U, 0U, 2U));
}

TEST(DepExtractor, DynamicEndpointOverWholeColumnStaysCompact) {
  const Workbook wb = Workbook::create();
  const ExtractedDeps deps = ExtractFrom("SUM(A1:INDEX(C:C,5))", wb);
  EXPECT_TRUE(HasRangeDep(deps, CellRangeDependency{0U, 0U, Sheet::kMaxRows - 1U, 0U, 2U}));
  EXPECT_LE(deps.cell_deps.size(), 1U);
}

TEST(DepExtractor, DynamicEndpointConditionStaysAPlainCell) {
  const Workbook wb = Workbook::create();
  std::vector<CellNodeId> expected = RectCells(0U, 0U, 4U, 0U, 3U);
  expected.push_back(CellNodeId{0U, 0U, 25U});  // Z1
  EXPECT_EQ(Sorted(ExtractFrom("SUM(A1:IF(Z1>0,B5,D5))", wb).cell_deps), Sorted(expected));
}

TEST(DepExtractor, OffsetEndpointKeepsDynamicReference) {
  const Workbook wb = Workbook::create();
  const ExtractedDeps deps = ExtractFrom("SUM(A1:OFFSET(A1,2,2))", wb);
  EXPECT_EQ(Sorted(deps.cell_deps), (std::vector<CellNodeId>{CellNodeId{0U, 0U, 0U}}));
  EXPECT_TRUE(deps.range_deps.empty());
  EXPECT_TRUE(deps.has_dynamic_reference);
}

TEST(DepExtractor, OffsetBaseIsNoRead) {
  // OFFSET's base only positions the rectangle it returns: `=OFFSET(D2,-1,0)`
  // in D2 reads D1, not D2, and is not circular (measured on Excel 365).
  // A reference-returning base only positions what it can return, so
  // `=OFFSET(INDEX(A1:A10,3),1,0)` in A3 is not circular either; the offsets
  // and the base's other arguments are still read.
  const Workbook wb = Workbook::create();
  const ExtractedDeps plain = ExtractFrom("OFFSET(A1,1,1)", wb);
  EXPECT_TRUE(plain.is_volatile);
  EXPECT_TRUE(plain.has_dynamic_reference);
  EXPECT_TRUE(plain.cell_deps.empty());
  EXPECT_EQ(Sorted(ExtractFrom("OFFSET(A1,B1,1)", wb).cell_deps), (std::vector<CellNodeId>{CellNodeId{0U, 0U, 1U}}));
  EXPECT_TRUE(ExtractFrom("OFFSET(INDEX(A1:A10,3),1,0)", wb).cell_deps.empty());
  EXPECT_EQ(Sorted(ExtractFrom("OFFSET(INDEX(A1:A10,B1),1,0)", wb).cell_deps),
            (std::vector<CellNodeId>{CellNodeId{0U, 0U, 1U}}));
}

TEST(DepExtractor, DynamicEndpointThroughLetBinding) {
  const Workbook wb = Workbook::create();
  EXPECT_EQ(Sorted(ExtractFrom("LET(r,C1:C10,SUM(A1:INDEX(r,3)))", wb).cell_deps), RectCells(0U, 0U, 9U, 0U, 2U));
}

TEST(DepExtractor, DynamicEndpointThroughDefinedName) {
  Workbook wb = Workbook::create();
  std::vector<DefinedName> names;
  names.push_back(DefinedName{"Rng", "=Sheet1!$C$1:$C$10", -1, false, ""});
  wb.set_defined_names(std::move(names));
  EXPECT_EQ(Sorted(ExtractFrom("SUM(A1:INDEX(Rng,3))", wb).cell_deps), RectCells(0U, 0U, 9U, 0U, 2U));
}

TEST(DepExtractor, DefinedNameEndpointRegistersBoundingBox) {
  // `A1:CellNm` reads A1:C3, including B2 which neither endpoint names.
  Workbook wb = Workbook::create();
  std::vector<DefinedName> names;
  names.push_back(DefinedName{"CellNm", "=Sheet1!$C$3", -1, false, ""});
  wb.set_defined_names(std::move(names));
  const std::vector<CellNodeId> deps = Sorted(ExtractFrom("SUM(A1:CellNm)", wb).cell_deps);
  EXPECT_EQ(deps, RectCells(0U, 0U, 2U, 0U, 2U));
}

TEST(DepExtractor, SelfBookNameReadsTheWorkbookScopedDefinition) {
  // `[0]!G` reads the workbook-scoped G (B2), never Sheet1's local G (C3).
  Workbook wb = Workbook::create();
  std::vector<DefinedName> names;
  names.push_back(DefinedName{"G", "=Sheet1!$C$3", 0, false, ""});
  names.push_back(DefinedName{"G", "=Sheet1!$B$2", -1, false, ""});
  names.push_back(DefinedName{"Rng", "=Sheet1!$A$10:$A$11", -1, false, ""});
  wb.set_defined_names(std::move(names));
  EXPECT_EQ(Sorted(ExtractFrom("[0]!G", wb).cell_deps), RectCells(0U, 1U, 1U, 1U, 1U));
  EXPECT_EQ(Sorted(ExtractFrom("SUM([0]!Rng)", wb).cell_deps), RectCells(0U, 9U, 10U, 0U, 0U));
  EXPECT_EQ(Sorted(ExtractFrom("SUM(A1:[0]!G)", wb).cell_deps), RectCells(0U, 0U, 1U, 0U, 1U));
}

TEST(DepExtractor, DynamicEndpointOnAnotherSheet) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  EXPECT_EQ(Sorted(ExtractFrom("SUM(Sheet2!A1:INDEX(Sheet2!C1:C3,2))", wb).cell_deps), RectCells(1U, 0U, 2U, 0U, 2U));
}

TEST(DepExtractor, ReferenceOnlyArgumentsRegisterNothing) {
  const Workbook wb = Workbook::create();
  for (const char* src :
       {"ROWS(A:A)", "ROW(A1)", "COLUMNS(1:1)", "ROWS(A1:C3)", "ISREF(A1)", "COLUMN(A1:B2:C3)", "AREAS((A1,B2:C3))"}) {
    const ExtractedDeps deps = ExtractFrom(src, wb);
    EXPECT_TRUE(deps.cell_deps.empty()) << src;
    EXPECT_TRUE(deps.range_deps.empty()) << src;
  }
}

TEST(DepExtractor, ReferenceOnlyArgumentThroughDefinedName) {
  Workbook wb = Workbook::create();
  std::vector<DefinedName> names;
  names.push_back(DefinedName{"Rng", "=Sheet1!$C$1:$C$10", -1, false, ""});
  wb.set_defined_names(std::move(names));
  const ExtractedDeps deps = ExtractFrom("ROWS(Rng)", wb);
  EXPECT_TRUE(deps.cell_deps.empty());
  EXPECT_TRUE(deps.range_deps.empty());
}

TEST(DepExtractor, ReferenceOnlyCallStillWalksComputedArguments) {
  const Workbook wb = Workbook::create();
  // A spill's extent is read from its anchor.
  EXPECT_EQ(ExtractFrom("ROWS(A1#)", wb).cell_deps, (std::vector<CellNodeId>{CellNodeId{0U, 0U, 0U}}));
  // INDEX reads B1 to pick the row.
  const std::vector<CellNodeId> index_deps = ExtractFrom("ROW(INDEX(A1:A10,B1))", wb).cell_deps;
  EXPECT_NE(std::find(index_deps.begin(), index_deps.end(), CellNodeId{0U, 0U, 1U}), index_deps.end());
}

TEST(DepExtractor, ShadowedReferenceOnlyNameKeepsItsArgument) {
  const Workbook wb = Workbook::create();
  EXPECT_EQ(Sorted(ExtractFrom("LET(rows,LAMBDA(x,x),rows(A1:A3))", wb).cell_deps), RectCells(0U, 0U, 2U, 0U, 0U));
}

TEST(DepExtractor, ThreeDSpanMetadataIsNormalizedAndDeduplicatedAcrossExpansion) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "First")));
  wb.add_sheet("Second");
  wb.add_sheet("Third");
  wb.set_defined_names({
      DefinedName{"Inner", "First:Second!A1:A3", -1, false, ""},
      DefinedName{"Outer", "Inner", -1, false, ""},
      DefinedName{"F", "LAMBDA(x,SUM(Outer)+x)", -1, false, ""},
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

}  // namespace
}  // namespace formulon::eval
