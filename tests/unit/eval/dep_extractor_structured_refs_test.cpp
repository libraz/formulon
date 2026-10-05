// Dependency extractor tests grouped by dependency source.

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
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

TableMetadata MakeTable(std::string name, std::string ref, std::size_t sheet_index, bool header_row, bool totals_row,
                        std::vector<std::string> column_names) {
  TableMetadata t;
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
    TableColumn col;
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
  std::vector<TableMetadata> tables;
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
  std::vector<TableMetadata> tables;
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
  std::vector<TableMetadata> tables;
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
  std::vector<TableMetadata> tables;
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
  std::vector<TableMetadata> tables;
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
  std::vector<TableMetadata> tables;
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
  std::vector<TableMetadata> tables;
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
  std::vector<TableMetadata> tables;
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
  std::vector<TableMetadata> tables;
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
  std::vector<TableMetadata> tables;
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
  std::vector<TableMetadata> tables;
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

}  // namespace
}  // namespace formulon::eval
