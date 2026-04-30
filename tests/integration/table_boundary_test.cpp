// Copyright 2026 libraz. Licensed under the MIT License.
//
// Integration tests for Excel structured (table) references at the
// boundary between tables and the rest of the workbook. These tests
// drive the full `Workbook::set_cell_*` -> `Workbook::recalc()` pipeline
// so the recalc engine, dep graph, structured-reference resolver, and
// evaluator all run end to end. They lock in current resolver behaviour
// at the table edges and across-table interaction surfaces.
//
// The contract under test:
//   * `Table[Col]` resolves to the column's data range and feeds range-
//     aware aggregators (`SUM` etc.) correctly.
//   * Two tables on the same sheet stay disjoint: `Alpha[Val]` does not
//     bleed into `Beta[Val]`.
//   * Cross-sheet references (`Sheet1!Sales[Amount]` shape) -- Excel does
//     NOT require the sheet qualifier because the table name is workbook-
//     unique; we verify that `=SUM(Sales[Amount])` typed on Sheet2
//     resolves to Sheet1's table data without any sheet prefix.
//   * Resizing the table (extending `ref` to cover more rows) updates
//     the resolved range, including totals/headers.
//   * `Table[@Col]` from a row OUTSIDE the data span surfaces #VALUE!.
//   * `Table[@Col]` from a row AT the header (`current_row == r_top`
//     when the header is present) surfaces #VALUE! because the header
//     row is not part of the data span.

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/tables_reader.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// Builds a single-table workbook anchored on Sheet1. The table covers
// A1:C5 with header row + 4 data rows (no totals row).
//
//   Row | A          B          C
//   ----+----------------------------------
//   1   | Region     Product    Amount   <- header row
//   2   | "North"    "Apple"     10
//   3   | "South"    "Banana"    20
//   4   | "East"     "Cherry"    30
//   5   | "West"     "Durian"    40
Workbook MakeSalesWorkbook() {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 0U, Value::text("Region"));
  s.set_cell_value(0U, 1U, Value::text("Product"));
  s.set_cell_value(0U, 2U, Value::text("Amount"));
  s.set_cell_value(1U, 0U, Value::text("North"));
  s.set_cell_value(1U, 1U, Value::text("Apple"));
  s.set_cell_value(1U, 2U, Value::number(10.0));
  s.set_cell_value(2U, 0U, Value::text("South"));
  s.set_cell_value(2U, 1U, Value::text("Banana"));
  s.set_cell_value(2U, 2U, Value::number(20.0));
  s.set_cell_value(3U, 0U, Value::text("East"));
  s.set_cell_value(3U, 1U, Value::text("Cherry"));
  s.set_cell_value(3U, 2U, Value::number(30.0));
  s.set_cell_value(4U, 0U, Value::text("West"));
  s.set_cell_value(4U, 1U, Value::text("Durian"));
  s.set_cell_value(4U, 2U, Value::number(40.0));

  io::TableMetadata table;
  table.id = 1U;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:C5";
  table.sheet_index = 0U;
  table.header_row = true;
  table.totals_row = false;
  table.columns = {
      io::TableColumn{1U, "Region", "", ""},
      io::TableColumn{2U, "Product", "", ""},
      io::TableColumn{3U, "Amount", "", ""},
  };
  std::vector<io::TableMetadata> tables = {std::move(table)};
  wb.set_tables(std::move(tables));
  return wb;
}

Value StoredValue(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& s = wb.sheet(sheet_index);
  if (const Cell* c = s.cell_at(row, col); c != nullptr) {
    return c->cached_value;
  }
  return Value::blank();
}

// ---------------------------------------------------------------------------
// Single-table aggregation
// ---------------------------------------------------------------------------

TEST(TableBoundary, SumOfColumnDataMatchesLiteralSum) {
  // E1 = =SUM(Sales[Amount]) -> 10 + 20 + 30 + 40 = 100.
  Workbook wb = MakeSalesWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 4U, "=SUM(Sales[Amount])")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value e1 = StoredValue(wb, 0U, 0U, 4U);
  ASSERT_TRUE(e1.is_number()) << "kind=" << static_cast<int>(e1.kind());
  EXPECT_DOUBLE_EQ(e1.as_number(), 100.0);
}

TEST(TableBoundary, ColumnRangeSpansMultipleColumns) {
  // E1 = =SUM(Sales[[Product]:[Amount]]) -> Product cells are text
  // (skipped by SUM); Amount cells contribute -> 100.
  Workbook wb = MakeSalesWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 4U, "=SUM(Sales[[Product]:[Amount]])")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value e1 = StoredValue(wb, 0U, 0U, 4U);
  ASSERT_TRUE(e1.is_number());
  EXPECT_DOUBLE_EQ(e1.as_number(), 100.0);
}

// ---------------------------------------------------------------------------
// Two tables on the same sheet stay disjoint
// ---------------------------------------------------------------------------

TEST(TableBoundary, TwoTablesSameSheetDoNotBleed) {
  // Alpha on A1:B3 (header + 2 data rows), columns Cat / Val.
  // Beta on D1:E3 (header + 2 data rows), columns Cat / Val.
  // Both tables have a column called "Val"; resolving Alpha[Val] must
  // NOT pull Beta's data and vice versa.
  //
  // Note on table names: a table identifier whose first 1-3 ASCII
  // characters could be column letters followed by ASCII digits (e.g.
  // `Tbl1` parses as cell ref TBL row 1) is recognised by the
  // tokenizer as a `CellRef`, not an `Ident`, and so the parser does
  // NOT treat the bracket payload as a structured reference. This is
  // current behaviour; Excel itself rejects table names that look like
  // cell references at table-creation time. We use `Alpha` / `Beta`
  // here to stay clear of the conflict.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 0U, Value::text("Cat"));
  s.set_cell_value(0U, 1U, Value::text("Val"));
  s.set_cell_value(1U, 0U, Value::text("a"));
  s.set_cell_value(1U, 1U, Value::number(1.0));
  s.set_cell_value(2U, 0U, Value::text("b"));
  s.set_cell_value(2U, 1U, Value::number(2.0));

  s.set_cell_value(0U, 3U, Value::text("Cat"));
  s.set_cell_value(0U, 4U, Value::text("Val"));
  s.set_cell_value(1U, 3U, Value::text("x"));
  s.set_cell_value(1U, 4U, Value::number(100.0));
  s.set_cell_value(2U, 3U, Value::text("y"));
  s.set_cell_value(2U, 4U, Value::number(200.0));

  io::TableMetadata t1;
  t1.id = 1U;
  t1.name = "Alpha";
  t1.display_name = "Alpha";
  t1.ref = "A1:B3";
  t1.sheet_index = 0U;
  t1.header_row = true;
  t1.totals_row = false;
  io::TableColumn t1_cat;
  t1_cat.id = 1U;
  t1_cat.name = "Cat";
  io::TableColumn t1_val;
  t1_val.id = 2U;
  t1_val.name = "Val";
  t1.columns = {t1_cat, t1_val};

  io::TableMetadata t2;
  t2.id = 2U;
  t2.name = "Beta";
  t2.display_name = "Beta";
  t2.ref = "D1:E3";
  t2.sheet_index = 0U;
  t2.header_row = true;
  t2.totals_row = false;
  io::TableColumn t2_cat;
  t2_cat.id = 1U;
  t2_cat.name = "Cat";
  io::TableColumn t2_val;
  t2_val.id = 2U;
  t2_val.name = "Val";
  t2.columns = {t2_cat, t2_val};

  std::vector<io::TableMetadata> tables = {std::move(t1), std::move(t2)};
  wb.set_tables(std::move(tables));

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 5U, 0U, "=SUM(Alpha[Val])")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 5U, 1U, "=SUM(Beta[Val])")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value a6 = StoredValue(wb, 0U, 5U, 0U);
  const Value b6 = StoredValue(wb, 0U, 5U, 1U);
  ASSERT_TRUE(a6.is_number()) << "Alpha[Val] expected 3.0; kind=" << static_cast<int>(a6.kind());
  ASSERT_TRUE(b6.is_number()) << "Beta[Val] expected 300.0; kind=" << static_cast<int>(b6.kind());
  EXPECT_DOUBLE_EQ(a6.as_number(), 3.0);
  EXPECT_DOUBLE_EQ(b6.as_number(), 300.0);
}

// ---------------------------------------------------------------------------
// Cross-sheet structured references
// ---------------------------------------------------------------------------

TEST(TableBoundary, StructuredRefFromAnotherSheetResolvesByName) {
  // Sheet1 holds the Sales table (A1:C5). Sheet2 has the formula
  // `=SUM(Sales[Amount])`. The structured-reference resolver looks up
  // the table by name in `wb.tables()` regardless of which sheet the
  // formula lives on, so the answer should be 100 even though no sheet
  // qualifier appears in the formula text.
  Workbook wb = MakeSalesWorkbook();
  wb.add_sheet("Sheet2");
  ASSERT_EQ(wb.sheet_count(), 2U);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 0U, 0U, "=SUM(Sales[Amount])")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value v = StoredValue(wb, 1U, 0U, 0U);
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_DOUBLE_EQ(v.as_number(), 100.0);
}

// ---------------------------------------------------------------------------
// `@`-row context (kThisRow) edge cases
// ---------------------------------------------------------------------------

TEST(TableBoundary, RowImplicitOutsideDataReturnsValueError) {
  // Formula in row 6 (0-based 5), outside the table's data span (rows
  // 1..4). Sales[@Amount] -> #VALUE!.
  Workbook wb = MakeSalesWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 5U, 4U, "=Sales[@Amount]")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value v = StoredValue(wb, 0U, 5U, 4U);
  ASSERT_TRUE(v.is_error()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TableBoundary, RowImplicitOnHeaderRowReturnsValueError) {
  // Formula in row 1 (0-based 0), which IS the table's header row.
  // The data span is rows 1..4 (0-based), so row 0 is outside ->
  // #VALUE!. This locks in the resolver's strict data-only behaviour
  // for `@`; Excel matches.
  Workbook wb = MakeSalesWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 4U, "=Sales[@Amount]")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value v = StoredValue(wb, 0U, 0U, 4U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TableBoundary, RowImplicitOnDataRowResolves) {
  // Formula in row 3 (0-based 2), which is the second data row ->
  // Sales[@Amount] picks the Amount cell of that row -> 20.
  Workbook wb = MakeSalesWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 4U, "=Sales[@Amount]")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value v = StoredValue(wb, 0U, 2U, 4U);
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);
}

// ---------------------------------------------------------------------------
// Resizing the table updates resolved ranges
// ---------------------------------------------------------------------------

TEST(TableBoundary, ResizingTableExtendsDataRange) {
  // Start with the 4-row table. Resize to A1:C7 (3 more data rows) and
  // populate the new rows. After re-running recalc, SUM(Sales[Amount])
  // should reflect the new total (10 + 20 + 30 + 40 + 50 + 60 + 70 = 280).
  Workbook wb = MakeSalesWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 4U, "=SUM(Sales[Amount])")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_DOUBLE_EQ(StoredValue(wb, 0U, 0U, 4U).as_number(), 100.0);

  // Add three more data rows.
  Sheet& s = wb.sheet(0);
  s.set_cell_value(5U, 0U, Value::text("North"));
  s.set_cell_value(5U, 1U, Value::text("Eggplant"));
  s.set_cell_value(5U, 2U, Value::number(50.0));
  s.set_cell_value(6U, 0U, Value::text("South"));
  s.set_cell_value(6U, 1U, Value::text("Fig"));
  s.set_cell_value(6U, 2U, Value::number(60.0));
  s.set_cell_value(7U, 0U, Value::text("East"));
  s.set_cell_value(7U, 1U, Value::text("Grape"));
  s.set_cell_value(7U, 2U, Value::number(70.0));

  // Re-emit the table metadata with the extended ref. (Workbook owns the
  // tables by value, so we replace the whole vector.) The resolver
  // re-reads `wb.tables()` on every evaluation, so the next recalc picks
  // up the new range.
  std::vector<io::TableMetadata> tables = wb.tables();
  ASSERT_EQ(tables.size(), 1U);
  tables[0].ref = "A1:C8";
  wb.set_tables(std::move(tables));

  // Re-mark the formula dirty so the recalc actually re-fires it. (No
  // dep edge currently connects table-metadata edits back to consumer
  // cells; future work will likely add an explicit "table reshape" hook
  // -- for now we simulate that by re-typing the formula.)
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 4U, "=SUM(Sales[Amount])")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value e1 = StoredValue(wb, 0U, 0U, 4U);
  ASSERT_TRUE(e1.is_number()) << "kind=" << static_cast<int>(e1.kind());
  EXPECT_DOUBLE_EQ(e1.as_number(), 280.0);
}

// ---------------------------------------------------------------------------
// Unknown-column / unknown-table negative paths
// ---------------------------------------------------------------------------

TEST(TableBoundary, UnknownTableNameSurfacesNameError) {
  // `=Other[Amount]` -> #NAME? (no table called Other).
  Workbook wb = MakeSalesWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 4U, "=Other[Amount]")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value v = StoredValue(wb, 0U, 0U, 4U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

}  // namespace
}  // namespace formulon
