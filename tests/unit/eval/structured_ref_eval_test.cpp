//
// End-to-end tests for structured (table) reference evaluation. Builds a
// workbook with a single registered table, populates the cells, and asserts
// each of the eight shapes Bundle 4.4 ships resolves to the expected value.

#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/structured_ref.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "table.h"
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

  TableMetadata table;
  table.id = 1;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = with_totals_row ? "A1:C5" : "A1:C4";
  table.sheet_index = 0;
  table.header_row = true;
  table.totals_row = with_totals_row;
  TableColumn region;
  region.id = 1;
  region.name = "Region";
  TableColumn product;
  product.id = 2;
  product.name = "Product";
  TableColumn amount;
  amount.id = 3;
  amount.name = "Amount";
  table.columns = {region, product, amount};
  std::vector<TableMetadata> tables = {table};
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
  EXPECT_EQ(sel.value().payload.specifiers, 0U);
  EXPECT_TRUE(sel.value().payload.columns.empty());
}

TEST(StructuredRefResolver, ParsePayloadSingleColumn) {
  auto sel = parse_structured_ref_payload("Region");
  ASSERT_TRUE(sel);
  EXPECT_EQ(sel.value().payload.specifiers, 0U);
  ASSERT_EQ(sel.value().payload.columns.size(), 1U);
  EXPECT_EQ(sel.value().payload.columns[0], "Region");
}

TEST(StructuredRefResolver, ParsePayloadAtColumn) {
  auto sel = parse_structured_ref_payload("@Amount");
  ASSERT_TRUE(sel);
  EXPECT_EQ(sel.value().payload.specifiers, StructuredRefSpecifiers::kThisRow);
  ASSERT_EQ(sel.value().payload.columns.size(), 1U);
  EXPECT_EQ(sel.value().payload.columns[0], "Amount");
}

TEST(StructuredRefResolver, ParsePayloadBareAtIsThisRow) {
  auto sel = parse_structured_ref_payload("@");
  ASSERT_TRUE(sel);
  EXPECT_EQ(sel.value().payload.specifiers, StructuredRefSpecifiers::kThisRow);
  EXPECT_TRUE(sel.value().payload.columns.empty());
}

TEST(StructuredRefResolver, ParsePayloadAllSpecifier) {
  auto sel = parse_structured_ref_payload("#All");
  ASSERT_TRUE(sel);
  EXPECT_EQ(sel.value().payload.specifiers, StructuredRefSpecifiers::kAll);
}

TEST(StructuredRefResolver, ParsePayloadHeadersAndColumn) {
  auto sel = parse_structured_ref_payload("[#Headers],[Region]");
  ASSERT_TRUE(sel);
  EXPECT_EQ(sel.value().payload.specifiers, StructuredRefSpecifiers::kHeaders);
  ASSERT_EQ(sel.value().payload.columns.size(), 1U);
  EXPECT_EQ(sel.value().payload.columns[0], "Region");
}

TEST(StructuredRefResolver, ParsePayloadColumnRange) {
  auto sel = parse_structured_ref_payload("[Region]:[Amount]");
  ASSERT_TRUE(sel);
  ASSERT_EQ(sel.value().payload.columns.size(), 2U);
  EXPECT_EQ(sel.value().payload.columns[0], "Region");
  EXPECT_EQ(sel.value().payload.columns[1], "Amount");
}

TEST(StructuredRefResolver, ParsePayloadUnbracketedCommaAndColonAreOneName) {
  auto comma = parse_structured_ref_payload("a,b");
  ASSERT_TRUE(comma);
  ASSERT_EQ(comma.value().payload.columns.size(), 1U);
  EXPECT_EQ(comma.value().payload.columns[0], "a,b");
  auto colon = parse_structured_ref_payload("a:b");
  ASSERT_TRUE(colon);
  ASSERT_EQ(colon.value().payload.columns.size(), 1U);
  EXPECT_EQ(colon.value().payload.columns[0], "a:b");
}

TEST(StructuredRefResolver, ParsePayloadUnknownSpecifierIsName) {
  auto sel = parse_structured_ref_payload("#Bogus");
  ASSERT_FALSE(sel);
  EXPECT_EQ(sel.error(), ErrorCode::Name);
}

TEST(StructuredRefResolver, OneRowTableIsRejectedBeforeTotalsRangeUnderflows) {
  Workbook wb = Workbook::create();
  TableMetadata table;
  table.id = 1;
  table.name = "Malformed";
  table.display_name = "Malformed";
  table.ref = "A1:A1";
  table.sheet_index = 0;
  table.header_row = true;
  table.totals_row = true;
  TableColumn column;
  column.id = 1;
  column.name = "Value";
  table.columns.push_back(std::move(column));
  wb.set_tables({std::move(table)});

  StructuredRefSelector selector;
  selector.table_name = "Malformed";
  auto resolved = resolve_structured_ref(selector, wb, 0U);
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

TEST(StructuredRefEval, SortProjectsTableBlankButNestedCountKeepsBlank) {
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  // Sales[Amount] is now {10, blank, 30}; the blank is an interior table
  // cell, so SORT keeps it as the last sort key and the grid boundary
  // projects it to numeric zero only for the direct result.
  wb.sheet(0).set_cell_value(2, 2, Value::blank());

  Arena direct_arena;
  const Value direct = EvalAt(wb, "=SORT(Sales[Amount])", /*formula_row=*/4, direct_arena);
  ASSERT_TRUE(direct.is_array()) << direct.debug_to_string();
  ASSERT_EQ(direct.as_array_rows(), 3U);
  ASSERT_EQ(direct.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(direct.as_array_cells()[0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(direct.as_array_cells()[1].as_number(), 30.0);
  ASSERT_TRUE(direct.as_array_cells()[2].is_number());
  EXPECT_DOUBLE_EQ(direct.as_array_cells()[2].as_number(), 0.0);

  Arena nested_arena;
  const Value count = EvalAt(wb, "=COUNT(SORT(Sales[Amount]))", /*formula_row=*/4, nested_arena);
  ASSERT_TRUE(count.is_number()) << count.debug_to_string();
  EXPECT_DOUBLE_EQ(count.as_number(), 2.0);
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

TEST(StructuredRefEval, EscapedApostropheColumnNameResolves) {
  // ECMA-376 18.5.1.10: a column name containing an apostrophe is spelled
  // with a doubled apostrophe inside the bracket (`Bob''s`), which must
  // unescape to `Bob's` before comparing against the table's real column
  // name -- not compared verbatim with the doubled spelling still in it.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::text("Bob's"));
  s.set_cell_value(1, 0, Value::number(42.0));
  TableMetadata table;
  table.id = 1;
  table.name = "Sales";
  table.ref = "A1:A2";
  table.sheet_index = 0;
  table.header_row = true;
  TableColumn col;
  col.id = 1;
  col.name = "Bob's";
  table.columns = {col};
  wb.set_tables({table});

  Arena arena;
  const Value v = EvalAt(wb, "=Sales[Bob''s]", /*formula_row=*/1, arena);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);
}

TEST(StructuredRefEval, EscapedSpecialCharacterColumnNamesResolve) {
  // Column names containing `#`, `[`, `]` and `'` are spelled with the
  // ECMA-376 18.5.1.10 apostrophe escape in formula text.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  const char* names[] = {"#Items", "Col[x]", "Bob's"};
  TableMetadata table;
  table.id = 1;
  table.name = "Table1";
  table.ref = "A1:C2";
  table.sheet_index = 0;
  table.header_row = true;
  for (std::uint32_t c = 0; c < 3; ++c) {
    s.set_cell_value(0, c, Value::text(names[c]));
    s.set_cell_value(1, c, Value::number(10.0 * (c + 1)));
    TableColumn col;
    col.id = c + 1;
    col.name = names[c];
    table.columns.push_back(col);
  }
  wb.set_tables({table});

  struct Case {
    const char* src;
    double want;
  };
  const Case cases[] = {
      {"=Table1['#Items]", 10.0},
      {"=Table1[Col'[x']]", 20.0},
      {"=Table1[[#This Row],[Bob''s]]", 30.0},
      {"=Table1[[#This Row],['#Items]]", 10.0},
      {"=Table1[@[Col'[x']]]", 20.0},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(c.src);
    Arena arena;
    const Value v = EvalAt(wb, c.src, /*formula_row=*/1, arena);
    ASSERT_TRUE(v.is_number());
    EXPECT_DOUBLE_EQ(v.as_number(), c.want);
  }
}

TEST(StructuredRefEval, UnbracketedCommaAndColonNamesAreOneColumn) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  const char* names[] = {"a,b", "a:b", "plain"};
  TableMetadata table;
  table.id = 1;
  table.name = "t";
  table.display_name = "t";
  table.ref = "A1:C2";
  table.sheet_index = 0;
  table.header_row = true;
  for (std::uint32_t c = 0; c < 3; ++c) {
    s.set_cell_value(0, c, Value::text(names[c]));
    s.set_cell_value(1, c, Value::number(10.0 * (c + 1)));
    TableColumn col;
    col.id = c + 1;
    col.name = names[c];
    table.columns.push_back(col);
  }
  wb.set_tables({table});

  struct Case {
    const char* src;
    double want;
  };
  const Case cases[] = {
      {"=SUM(t[a,b])", 10.0},
      {"=SUM(t[a:b])", 20.0},
      {"=t[[#This Row],[a:b]]", 20.0},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(c.src);
    Arena arena;
    const Value v = EvalAt(wb, c.src, /*formula_row=*/1, arena);
    ASSERT_TRUE(v.is_number()) << v.debug_to_string();
    EXPECT_DOUBLE_EQ(v.as_number(), c.want);
  }
}

TEST(StructuredRefEval, BareAtResolvesToThisRow) {
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  Arena arena;
  const Value v = EvalAt(wb, "=SUM(Sales[@])", /*formula_row=*/2, arena);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_EQ(v.as_number(), 20.0);
}

TEST(StructuredRefEval, StoredDoubledApostropheColumnResolves) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::text("a'b"));
  s.set_cell_value(1, 0, Value::number(7.0));
  TableMetadata table;
  table.id = 1;
  table.name = "t";
  table.display_name = "t";
  table.ref = "A1:A2";
  table.sheet_index = 0;
  table.header_row = true;
  TableColumn col;
  col.id = 1;
  col.name = "a'b";
  table.columns = {col};
  wb.set_tables({table});

  Arena arena;
  const Value v = EvalAt(wb, "=t[a''b]", /*formula_row=*/1, arena);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

TEST(StructuredRefEval, HeadersOnTableWithoutHeadersIsRef) {
  // Build a table with header_row=false so `[#Headers]` produces #REF!.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(1.0));
  s.set_cell_value(1, 0, Value::number(2.0));
  TableMetadata table;
  table.id = 1;
  table.name = "NoHead";
  table.display_name = "NoHead";
  table.ref = "A1:A2";
  table.sheet_index = 0;
  table.header_row = false;
  table.totals_row = false;
  TableColumn c;
  c.id = 1;
  c.name = "Col1";
  table.columns = {c};
  std::vector<TableMetadata> tables = {table};
  wb.set_tables(std::move(tables));

  Arena arena;
  const Value v = EvalAt(wb, "=NoHead[[#Headers],[Col1]]", /*formula_row=*/2, arena);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// Evaluates `form` with `<T>` replaced by `ref` from (`row`, D).
std::string EvalTwin(const Workbook& wb, std::string form, std::string_view ref, std::uint32_t row) {
  for (std::size_t at = form.find("<T>"); at != std::string::npos; at = form.find("<T>", at + ref.size())) {
    form.replace(at, 3, ref);
  }
  Arena arena;
  return EvalAt(wb, "=" + form, row, arena).debug_to_string();
}

struct TwinCase {
  const char* form;
  Value expected;
};

// A table reference is a reference wherever the area it denotes is one: each form reads the same through
// `Sales[Amount]` as through `$C$2:$C$4`.
TEST(StructuredRefEval, TableReferenceIsTheAreaItDenotes) {
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  const TwinCase cases[] = {
      {"@<T>", Value::number(20.0)},
      {"@(<T>)", Value::number(20.0)},
      {"LET(a,<T>,@a)", Value::number(20.0)},
      {"@INDEX(<T>,0)", Value::number(20.0)},
      {"@INDEX(<T>,0,1)", Value::number(20.0)},
      {"@OFFSET(<T>,0,0)", Value::number(20.0)},
      {"@IF(TRUE,<T>)", Value::number(20.0)},
      {"@CHOOSE(1,<T>)", Value::number(20.0)},
      {"@TAKE(<T>,2)", Value::number(20.0)},
      {"ISREF(<T>)", Value::boolean(true)},
      {"ISREF((<T>))", Value::boolean(true)},
      {"ISREF(INDEX(<T>,1))", Value::boolean(true)},
      {"ISREF(TAKE(<T>,1))", Value::boolean(true)},
      {"ISREF(IF(TRUE,<T>))", Value::boolean(true)},
      {"AND(MAP(<T>,LAMBDA(x,ISREF(x))))", Value::boolean(true)},
      {"SUM(MAP(<T>,LAMBDA(x,ROW(x))))", Value::number(9.0)},
      {"ROWS(<T>)", Value::number(3.0)},
      {"ROWS(<T>:C9)", Value::number(8.0)},
      {"AREAS(<T>)", Value::number(1.0)},
      {"SUM(OFFSET(<T>,1,0))", Value::number(50.0)},
      {"CELL(\"address\",<T>)", Value::text("$C$2")},
      {"COLUMN(<T>)", Value::number(3.0)},
      {"SUM(ROW(<T>))", Value::number(9.0)},
      {"SHEET(<T>)", Value::number(1.0)},
      {"N(<T>)", Value::number(10.0)},
      {"OR(ISFORMULA(<T>))", Value::boolean(false)},
      {"SUM(--ISNA(FORMULATEXT(<T>)))", Value::number(3.0)},
  };
  for (const TwinCase& c : cases) {
    const std::string table = EvalTwin(wb, c.form, "Sales[Amount]", /*row=*/2);
    EXPECT_EQ(table, EvalTwin(wb, c.form, "$C$2:$C$4", /*row=*/2)) << c.form;
    EXPECT_EQ(table, c.expected.debug_to_string()) << c.form;
  }
}

// `@` over a table reference intersects the formula row like `@` over its area: a row beside the column reads that
// row's cell, the header row and a row below the table have no intersection.
TEST(StructuredRefEval, AtOverTableReferenceIntersectsTheFormulaRow) {
  Workbook wb = MakeWorkbookWithSalesTable(/*with_totals_row=*/false);
  const char* const forms[] = {"@<T>", "LET(a,<T>,@a)", "@OFFSET(<T>,0,0)", "@TAKE(<T>,3)"};
  for (const char* form : forms) {
    for (const std::uint32_t row : {0U, 6U}) {
      const std::string table = EvalTwin(wb, form, "Sales[Amount]", row);
      EXPECT_EQ(table, EvalTwin(wb, form, "$C$2:$C$4", row)) << form << " row " << row;
      EXPECT_EQ(table, Value::error(ErrorCode::Value).debug_to_string()) << form << " row " << row;
    }
    EXPECT_EQ(EvalTwin(wb, form, "Sales[Amount]", /*row=*/3U), Value::number(30.0).debug_to_string()) << form;
  }
}

// Table1 on A1:B4 (Qty, Price) with Qty = 1, 2, 3 in A2:A4.
Workbook MakeWorkbookWithQtyTable() {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::text("Qty"));
  s.set_cell_value(0, 1, Value::text("Price"));
  for (std::uint32_t r = 1; r <= 3; ++r) {
    s.set_cell_value(r, 0, Value::number(static_cast<double>(r)));
  }
  TableMetadata table;
  table.id = 1;
  table.name = "Table1";
  table.display_name = "Table1";
  table.ref = "A1:B4";
  table.sheet_index = 0;
  table.header_row = true;
  TableColumn qty;
  qty.id = 1;
  qty.name = "Qty";
  TableColumn price;
  price.id = 2;
  price.name = "Price";
  table.columns = {qty, price};
  wb.set_tables({table});
  return wb;
}

// The value as a cell shows it, a spilled array as its elements joined by `|`.
std::string Shown(const Value& v) {
  if (v.is_array()) {
    std::string out;
    for (std::size_t i = 0; i < static_cast<std::size_t>(v.as_array_rows()) * v.as_array_cols(); ++i) {
      out += (i == 0U ? "" : "|") + Shown(v.as_array_cells()[i]);
    }
    return out;
  }
  if (v.is_error()) {
    return display_name(v.as_error());
  }
  if (v.is_boolean()) {
    return v.as_boolean() ? "TRUE" : "FALSE";
  }
  if (v.is_number()) {
    char buf[32];
    std::snprintf(buf, sizeof(buf), "%g", v.as_number());
    return buf;
  }
  return v.is_text() ? std::string(v.as_text()) : v.debug_to_string();
}

std::string EvalShownAt(const Workbook& wb, std::string form, std::string_view ref, std::uint32_t row,
                        std::uint32_t col) {
  for (std::size_t at = form.find("<X>"); at != std::string::npos; at = form.find("<X>", at + ref.size())) {
    form.replace(at, 3, ref);
  }
  // The parser keeps views into its source, so the text outlives it.
  const std::string source = "=" + form;
  Arena arena;
  auto p = parser::Parser(source, arena);
  parser::AstNode* root = p.parse();
  if (root == nullptr) {
    return "PARSE";
  }
  EvalState state;
  EvalContext ctx(wb, wb.sheet(0), state);
  ctx = ctx.with_formula_cell(row, col);
  return Shown(evaluate(*root, arena, default_registry(), ctx));
}

// Reference functions over `@X`, `LET(a,X,a)`, a nested LET and a LET shadowed by a value, from E3 (beside data row
// 3) and A7 (below the table), with X the area A2:A4 and the table column Table1[Qty]. `@X` is the intersected cell
// beside the table and #VALUE! passed in below it, not a reference to nothing.
TEST(StructuredRefEval, ReferenceFunctionsOverAtAndLetWrappers) {
  const Workbook wb = MakeWorkbookWithQtyTable();
  const char* const wrappers[] = {"@<X>", "LET(a,<X>,a)", "LET(a,<X>,LET(b,a,b))", "LET(a,<X>,LET(a,1,a))"};
  struct Row {
    const char* fn;
    const char* shown[4][2];  // [wrapper][E3, A7]
  };
  const Row rows[] = {
      {"ROW(<W>)", {{"3", "#VALUE!"}, {"2|3|4", "2|3|4"}, {"2|3|4", "2|3|4"}, {"#VALUE!", "#VALUE!"}}},
      {"COLUMN(<W>)", {{"1", "#VALUE!"}, {"1", "1"}, {"1", "1"}, {"#VALUE!", "#VALUE!"}}},
      {"ROWS(<W>)", {{"1", "#VALUE!"}, {"3", "3"}, {"3", "3"}, {"1", "1"}}},
      {"COLUMNS(<W>)", {{"1", "#VALUE!"}, {"1", "1"}, {"1", "1"}, {"1", "1"}}},
      {"AREAS(<W>)", {{"1", "#VALUE!"}, {"1", "1"}, {"1", "1"}, {"#VALUE!", "#VALUE!"}}},
      {"ISREF(<W>)", {{"TRUE", "FALSE"}, {"TRUE", "TRUE"}, {"TRUE", "TRUE"}, {"FALSE", "FALSE"}}},
      {"SUM(OFFSET(<W>,0,0))", {{"2", "#VALUE!"}, {"6", "6"}, {"6", "6"}, {"#VALUE!", "#VALUE!"}}},
      {"CELL(\"address\",<W>)", {{"$A$3", "#VALUE!"}, {"$A$2", "$A$2"}, {"$A$2", "$A$2"}, {"#VALUE!", "#VALUE!"}}},
      {"SHEET(<W>)", {{"1", "#VALUE!"}, {"1", "1"}, {"1", "1"}, {"#N/A", "#N/A"}}},
      {"FORMULATEXT(<W>)",
       {{"#N/A", "#VALUE!"},
        {"#N/A|#N/A|#N/A", "#N/A|#N/A|#N/A"},
        {"#N/A|#N/A|#N/A", "#N/A|#N/A|#N/A"},
        {"#VALUE!", "#VALUE!"}}},
      {"ISFORMULA(<W>)",
       {{"FALSE", "#VALUE!"},
        {"FALSE|FALSE|FALSE", "FALSE|FALSE|FALSE"},
        {"FALSE|FALSE|FALSE", "FALSE|FALSE|FALSE"},
        {"#VALUE!", "#VALUE!"}}},
      {"INDEX(<W>,1,1)", {{"2", "#VALUE!"}, {"1", "1"}, {"1", "1"}, {"1", "1"}}},
  };
  for (const char* x : {"A2:A4", "Table1[Qty]"}) {
    for (const Row& row : rows) {
      for (std::size_t w = 0; w < 4U; ++w) {
        std::string form = row.fn;
        form.replace(form.find("<W>"), 3, wrappers[w]);
        EXPECT_EQ(EvalShownAt(wb, form, x, /*row=*/2U, /*col=*/4U), row.shown[w][0]) << form << " X=" << x << " E3";
        EXPECT_EQ(EvalShownAt(wb, form, x, /*row=*/6U, /*col=*/0U), row.shown[w][1]) << form << " X=" << x << " A7";
      }
    }
  }
}

// `@` takes a chain of cell references whole; a LET, a lambda call or an `@` standing for a reference closes a range.
TEST(StructuredRefEval, RangeEndpointsOverAtAndLet) {
  const Workbook wb = MakeWorkbookWithQtyTable();
  EXPECT_EQ(EvalShownAt(wb, "ROWS(@A2:A4:A9)", "", 2U, 4U), "1");
  EXPECT_EQ(EvalShownAt(wb, "ROWS(@A2:A4:A9)", "", 6U, 0U), "1");
  for (const char* x : {"A2:A4", "Table1[Qty]"}) {
    EXPECT_EQ(EvalShownAt(wb, "ROWS(LET(a,<X>,a):A9)", x, 2U, 4U), "8") << x;
    EXPECT_EQ(EvalShownAt(wb, "ROWS(LET(a,<X>,a):A9)", x, 6U, 0U), "8") << x;
  }
  EXPECT_EQ(EvalShownAt(wb, "ROWS(A9:@A2:A4)", "", 2U, 4U), "7");
  EXPECT_EQ(EvalShownAt(wb, "ROWS(A9:@A2:A4)", "", 6U, 0U), "#VALUE!");
  EXPECT_EQ(EvalShownAt(wb, "ROWS(A9:LAMBDA(x,x)(A2:A4))", "", 2U, 4U), "8");
  // At the top level each range spills the area A2:A9, its blank cells as 0.
  for (const char* form :
       {"LET(a,A2:A4,a):A9", "LET(a,Table1[Qty],a):A9", "A9:LET(a,A2:A4,a)", "LAMBDA(x,x)(A2:A4):A9"}) {
    EXPECT_EQ(EvalShownAt(wb, form, "", 2U, 4U), "1|2|3|0|0|0|0|0") << form;
  }
}

// A leading `@` takes the whole `:` chain and a trailing one only its operand, typed or as stored (`SINGLE`).
TEST(StructuredRefEval, AtOverTableChainTakesTheChainOrTheOperand) {
  const Workbook wb = MakeWorkbookWithQtyTable();
  for (const char* at : {"@Table1[Qty]:A9", "_xlfn.SINGLE(Table1[Qty]:A9)"}) {
    const std::string chain(at);
    const std::string rows = "ROWS(" + chain + ")";
    EXPECT_EQ(EvalShownAt(wb, rows, "", 2U, 4U), "1") << chain;
    EXPECT_EQ(EvalShownAt(wb, rows, "", 7U, 4U), "1") << chain;
    if (chain[0] == '@') {
      EXPECT_EQ(EvalShownAt(wb, rows, "", 214U, 0U), "#VALUE!") << chain;
      EXPECT_EQ(EvalShownAt(wb, chain, "", 2U, 4U), "2") << chain;
    }
    EXPECT_EQ(EvalShownAt(wb, "CELL(\"address\"," + chain + ")", "", 2U, 4U), "$A$3") << chain;
  }
  for (const char* tail : {"A9:@Table1[Qty]", "A9:_xlfn.SINGLE(Table1[Qty])"}) {
    const std::string rows = std::string("ROWS(") + tail + ")";
    EXPECT_EQ(EvalShownAt(wb, rows, "", 2U, 4U), "7") << tail;
    EXPECT_EQ(EvalShownAt(wb, rows, "", 218U, 0U), "#VALUE!") << tail;
  }
}

// `@` over a chain ending in a computed reference intersects the area it evaluates to, as over any multi-cell area.
TEST(StructuredRefEval, AtOverAComputedChainIntersectsItsArea) {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5U; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(10.0 * (r + 1U)));
  }
  for (const char* at : {"@A1:INDEX(A1:A5,3)", "_xlfn.SINGLE(A1:INDEX(A1:A5,3))"}) {
    const std::string chain(at);
    EXPECT_EQ(EvalShownAt(wb, "ROW(" + chain + ")", "", 1U, 4U), "2") << chain;
    EXPECT_EQ(EvalShownAt(wb, "ROW(" + chain + ")", "", 4U, 4U), "#VALUE!") << chain;
    EXPECT_EQ(EvalShownAt(wb, "ROWS(" + chain + ")", "", 1U, 4U), "1") << chain;
    EXPECT_EQ(EvalShownAt(wb, "ROWS(" + chain + ")", "", 4U, 4U), "#VALUE!") << chain;
    EXPECT_EQ(EvalShownAt(wb, chain, "", 1U, 4U), "20") << chain;
    EXPECT_EQ(EvalShownAt(wb, chain, "", 4U, 4U), "#VALUE!") << chain;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
