//
// End-to-end tests for structured (table) reference evaluation. Builds a
// workbook with a single registered table, populates the cells, and asserts
// each of the eight shapes Bundle 4.4 ships resolves to the expected value.

#include <cstdint>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/structured_ref.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "io/tables_reader.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Builds a workbook with one sheet "Sheet1" and one table "Sales" anchored at
// A1:C5 (header row + 3 data rows + totals row). Columns are Region / Product
// / Amount. Cells:
//
//   Row | A           B          C
//   ----+----------------------------------
//   1   | Region      Product    Amount   <- header row
//   2   | "North"     "Apple"     10
//   3   | "South"     "Banana"    20
//   4   | "East"      "Cherry"    30
//   5   | "Total"      ""        60       <- totals row
Workbook MakeWorkbookWithSalesTable(bool with_totals_row) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::text("Region"));
  s.set_cell_value(0, 1, Value::text("Product"));
  s.set_cell_value(0, 2, Value::text("Amount"));
  s.set_cell_value(1, 0, Value::text("North"));
  s.set_cell_value(1, 1, Value::text("Apple"));
  s.set_cell_value(1, 2, Value::number(10.0));
  s.set_cell_value(2, 0, Value::text("South"));
  s.set_cell_value(2, 1, Value::text("Banana"));
  s.set_cell_value(2, 2, Value::number(20.0));
  s.set_cell_value(3, 0, Value::text("East"));
  s.set_cell_value(3, 1, Value::text("Cherry"));
  s.set_cell_value(3, 2, Value::number(30.0));
  if (with_totals_row) {
    s.set_cell_value(4, 0, Value::text("Total"));
    s.set_cell_value(4, 1, Value::text(""));
    s.set_cell_value(4, 2, Value::number(60.0));
  }

  io::TableMetadata table;
  table.id = 1;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = with_totals_row ? "A1:C5" : "A1:C4";
  table.sheet_index = 0;
  table.header_row = true;
  table.totals_row = with_totals_row;
  io::TableColumn region;
  region.id = 1;
  region.name = "Region";
  io::TableColumn product;
  product.id = 2;
  product.name = "Product";
  io::TableColumn amount;
  amount.id = 3;
  amount.name = "Amount";
  table.columns = {region, product, amount};
  std::vector<io::TableMetadata> tables = {table};
  wb.set_tables(std::move(tables));
  return wb;
}

// Parses `formula`, evaluates against the workbook's first sheet anchored at
// `formula_row`, and returns the result. The caller's arena must outlive the
// returned `Value`.
Value EvalAt(const Workbook& wb, std::string_view formula, std::uint32_t formula_row, Arena& arena) {
  auto p = parser::Parser(formula, arena);
  parser::AstNode* root = p.parse();
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  EvalState state;
  EvalContext ctx(wb, wb.sheet(0), state);
  ctx = ctx.with_formula_cell(formula_row, /*col=*/3);
  return evaluate(*root, arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// Resolver-level tests (no parser involved)
// ---------------------------------------------------------------------------

TEST(StructuredRefResolver, ParsePayloadEmptyIsData) {
  auto sel = parse_structured_ref_payload("");
  ASSERT_TRUE(sel);
  EXPECT_EQ(sel.value().specifiers, StructuredRefSpecifiers::kData);
  EXPECT_FALSE(sel.value().column_first_set);
}

TEST(StructuredRefResolver, ParsePayloadSingleColumn) {
  auto sel = parse_structured_ref_payload("Region");
  ASSERT_TRUE(sel);
  EXPECT_EQ(sel.value().specifiers, StructuredRefSpecifiers::kData);
  ASSERT_TRUE(sel.value().column_first_set);
  EXPECT_EQ(sel.value().column_first, "Region");
  EXPECT_FALSE(sel.value().column_last_set);
}

TEST(StructuredRefResolver, ParsePayloadAtColumn) {
  auto sel = parse_structured_ref_payload("@Amount");
  ASSERT_TRUE(sel);
  EXPECT_EQ(sel.value().specifiers, StructuredRefSpecifiers::kThisRow | StructuredRefSpecifiers::kData);
  ASSERT_TRUE(sel.value().column_first_set);
  EXPECT_EQ(sel.value().column_first, "Amount");
}

TEST(StructuredRefResolver, ParsePayloadAllSpecifier) {
  auto sel = parse_structured_ref_payload("#All");
  ASSERT_TRUE(sel);
  EXPECT_EQ(sel.value().specifiers, StructuredRefSpecifiers::kAll);
}

TEST(StructuredRefResolver, ParsePayloadHeadersAndColumn) {
  auto sel = parse_structured_ref_payload("[#Headers],[Region]");
  ASSERT_TRUE(sel);
  EXPECT_EQ(sel.value().specifiers, StructuredRefSpecifiers::kHeaders);
  ASSERT_TRUE(sel.value().column_first_set);
  EXPECT_EQ(sel.value().column_first, "Region");
}

TEST(StructuredRefResolver, ParsePayloadColumnRange) {
  auto sel = parse_structured_ref_payload("[Region]:[Amount]");
  ASSERT_TRUE(sel);
  ASSERT_TRUE(sel.value().column_first_set);
  ASSERT_TRUE(sel.value().column_last_set);
  EXPECT_EQ(sel.value().column_first, "Region");
  EXPECT_EQ(sel.value().column_last, "Amount");
}

TEST(StructuredRefResolver, ParsePayloadUnknownSpecifierIsName) {
  auto sel = parse_structured_ref_payload("#Bogus");
  ASSERT_FALSE(sel);
  EXPECT_EQ(sel.error(), ErrorCode::Name);
}

TEST(StructuredRefResolver, OneRowTableIsRejectedBeforeTotalsRangeUnderflows) {
  Workbook wb = Workbook::create();
  io::TableMetadata table;
  table.id = 1;
  table.name = "Malformed";
  table.display_name = "Malformed";
  table.ref = "A1:A1";
  table.sheet_index = 0;
  table.header_row = true;
  table.totals_row = true;
  io::TableColumn column;
  column.id = 1;
  column.name = "Value";
  table.columns.push_back(std::move(column));
  wb.set_tables({std::move(table)});

  StructuredRefSelector selector;
  selector.table_name = "Malformed";
  auto resolved = resolve_structured_ref(selector, wb, 0U, 0U);
  ASSERT_FALSE(resolved);
  EXPECT_EQ(resolved.error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// End-to-end evaluation: each of the 8 shapes from Bundle 4.4
// ---------------------------------------------------------------------------

TEST(StructuredRefEval, SingleColumnDataReturnsArray) {
  // `Sales[Amount]` -> the 3-cell data range C2:C4.
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  Arena arena;
  // Wrap in SUM so the test exercises range-aware aggregation.
  const Value v = EvalAt(wb, "=SUM(Sales[Amount])", /*formula_row=*/4, arena);
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_number(), 60.0);
}

TEST(StructuredRefEval, RowImplicitOnDataRow) {
  // `Sales[@Amount]` from row index 2 (cell D3) -> should pick row 3 -> 20.
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  Arena arena;
  const Value v = EvalAt(wb, "=Sales[@Amount]", /*formula_row=*/2, arena);
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_number(), 20.0);
}

TEST(StructuredRefEval, RowImplicitOutsideDataIsValue) {
  // Formula in row 0 (header row) is outside the data span -> #VALUE!.
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  Arena arena;
  const Value v = EvalAt(wb, "=Sales[@Amount]", /*formula_row=*/0, arena);
  ASSERT_TRUE(v.is_error()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(StructuredRefEval, AllSpansEntireTable) {
  // `SUM(Sales[#All])` -> 10 + 20 + 30 = 60 (text cells coerce to skip in SUM).
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  Arena arena;
  const Value v = EvalAt(wb, "=SUM(Sales[#All])", /*formula_row=*/4, arena);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 60.0);
}

TEST(StructuredRefEval, HeadersReturnsHeaderRow) {
  // `Sales[[#Headers],[Amount]]` -> the single cell C1 holding "Amount".
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  Arena arena;
  const Value v = EvalAt(wb, "=Sales[[#Headers],[Amount]]", /*formula_row=*/4, arena);
  ASSERT_TRUE(v.is_text()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_text(), "Amount");
}

TEST(StructuredRefEval, TotalsReturnsTotalsRow) {
  // `Sales[[#Totals],[Amount]]` -> single cell C5 holding 60.
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/true);
  Arena arena;
  const Value v = EvalAt(wb, "=Sales[[#Totals],[Amount]]", /*formula_row=*/5, arena);
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_number(), 60.0);
}

TEST(StructuredRefEval, TotalsOnTableWithoutTotalsIsRef) {
  // No totals row -> #REF! on `[#Totals]`.
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  Arena arena;
  const Value v = EvalAt(wb, "=Sales[[#Totals],[Amount]]", /*formula_row=*/4, arena);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(StructuredRefEval, DataExplicitMatchesDefault) {
  // `SUM(Sales[[#Data],[Amount]])` -> same as `SUM(Sales[Amount])`.
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/true);
  Arena arena;
  const Value v = EvalAt(wb, "=SUM(Sales[[#Data],[Amount]])", /*formula_row=*/5, arena);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 60.0);
}

TEST(StructuredRefEval, ColumnRangeAcrossMultipleColumns) {
  // `SUM(Sales[[Product]:[Amount]])` -> sum over data rows of Product+Amount
  // columns. Product cells are text (skipped); Amount cells contribute.
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  Arena arena;
  const Value v = EvalAt(wb, "=SUM(Sales[[Product]:[Amount]])", /*formula_row=*/4, arena);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 60.0);
}

TEST(StructuredRefEval, MultiSpecifierHeadersAndColumn) {
  // `Sales[[#Headers],[Region]]` -> "Region" header text.
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  Arena arena;
  const Value v = EvalAt(wb, "=Sales[[#Headers],[Region]]", /*formula_row=*/4, arena);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "Region");
}

// ---------------------------------------------------------------------------
// Error edges
// ---------------------------------------------------------------------------

TEST(StructuredRefEval, MissingTableIsName) {
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  Arena arena;
  const Value v = EvalAt(wb, "=Other[Amount]", /*formula_row=*/4, arena);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(StructuredRefEval, UnknownColumnIsRef) {
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  Arena arena;
  const Value v = EvalAt(wb, "=Sales[Bogus]", /*formula_row=*/4, arena);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(StructuredRefEval, HeadersOnTableWithoutHeadersIsRef) {
  // Build a table with header_row=false so `[#Headers]` produces #REF!.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(1.0));
  s.set_cell_value(1, 0, Value::number(2.0));
  io::TableMetadata table;
  table.id = 1;
  table.name = "NoHead";
  table.display_name = "NoHead";
  table.ref = "A1:A2";
  table.sheet_index = 0;
  table.header_row = false;
  table.totals_row = false;
  io::TableColumn c;
  c.id = 1;
  c.name = "Col1";
  table.columns = {c};
  std::vector<io::TableMetadata> tables = {table};
  wb.set_tables(std::move(tables));

  Arena arena;
  const Value v = EvalAt(wb, "=NoHead[[#Headers],[Col1]]", /*formula_row=*/2, arena);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
