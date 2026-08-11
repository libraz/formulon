//
// Tests for Excel 365's PIVOTBY dynamic-array function. PIVOTBY shares the
// lazy-dispatch entry point and aggregator-resolution machinery with
// GROUPBY (see `tests/unit/eval/builtins_groupby_test.cpp`). The goal of
// this suite is to cover PIVOTBY's incremental surface:
//
//   * Form A / B / C aggregator resolution (parity with GROUPBY).
//   * 2D output layout: row-axis groups + col-axis groups + per-cell body.
//   * `field_headers` ∈ {0, 1, 2, 3} drives both row and col header edges.
//   * `row_total_depth` / `col_total_depth` ∈ {-1, 0, 1} control row /
//     column total placement (top/bottom for row totals, left/right for
//     column totals).
//   * `row_sort_order` / `col_sort_order` ∈ {-1, 0, 1} sort by row / col
//     totals (the "first/only value column" reduces to the total in the
//     single-column-values scope of this commit).
//   * `filter_array` masks data rows.
//   * Per-(row, col) error isolation: one cell errors, the rest succeed.
//   * ja-JP `fold_jp_text` applies symmetrically to row keys and col keys.
//   * Mismatched row counts -> `#VALUE!`.
//   * Multi-column row_fields / col_fields / values produce the K x L x V
//     cross-product layout (see `eval_pivotby_lazy` source comment for
//     the row / column ordering and the documented Mac Excel divergences).
//   * Empty / fully-excluded input -> `#CALC!`.
//   * Bad arg domains -> `#VALUE!`.
//   * Aggregator returning array -> `#CALC!` for that cell.

#include <cstdint>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it through the default function registry.
// Parse / eval arenas have process lifetime via `static thread_local` so the
// returned `Value`'s arena-backed payload (text, array cells, lambda body)
// survives until the next `EvalSrc` call. Each call resets the shared arenas
// to avoid cross-test contamination, matching the helper pattern in
// `builtins_countif_test.cpp`. A stack-local `ParseAndEval` would free both
// arenas at function return, leaving the returned `Value` dangling — caught
// by AddressSanitizer as heap-use-after-free in `Value::as_array_cells()`.
Value EvalSrc(std::string_view src) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_TRUE(p.errors().empty()) << "unexpected parse errors for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena, default_registry(), test::mac_context());
}

// Reads cell `(r, c)` from a 2D ArrayValue using row-major indexing.
const Value& Cell(const Value& v, std::uint32_t r, std::uint32_t c) {
  return v.as_array_cells()[static_cast<std::size_t>(r) * v.as_array_cols() + c];
}

// ---------------------------------------------------------------------------
// Form A: inline LAMBDA literal aggregator
// ---------------------------------------------------------------------------

TEST(PivotBy, BasicTwoByTwoLambdaAggregatorSum) {
  // Row groups: A, B. Col groups: X, Y.
  // Body: (A,X)=1, (A,Y)=2, (B,X)=3, (B,Y)=4.
  // With default field_headers=3, row_total_depth=1, col_total_depth=1:
  //   header row, 2 body rows, then bottom grand-total row.
  const Value v = EvalSrc(
      "=PIVOTBY({\"R\";\"A\";\"A\";\"B\";\"B\"}, {\"C\";\"X\";\"Y\";\"X\";\"Y\"},"
      "         {\"V\";1;2;3;4}, LAMBDA(v, SUM(v)))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Layout: header row + 2 body rows + bottom grand-total row = 4 rows.
  // Cols: row label + 2 col keys + grand-total = 4 cols.
  EXPECT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  // Header row: ["R", "X", "Y", "合計"].
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "R");
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "合計");
  // A row: [A, 1, 2, 3].
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 3).as_number(), 3.0);
  // B row: [B, 3, 4, 7].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 3.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 7.0);
  // Bottom grand total row: ["合計", X-total=4, Y-total=6, total=10].
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 3, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 6.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 3).as_number(), 10.0);
}

// ---------------------------------------------------------------------------
// Form C: bare function name aggregator
// ---------------------------------------------------------------------------

TEST(PivotBy, BareSumName) {
  // No headers (field_headers=0), no totals (row_total_depth=0,
  // col_total_depth=0). Pure 2x2 body.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 2).as_number(), 2.0);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 3.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 4.0);
}

TEST(PivotBy, BareSumFiltersRangeSourcedNonNumbers) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"X\";\"X\";\"X\"},"
      "         {1;TRUE;\"text\";2}, SUM, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 2.0);
}

TEST(PivotBy, BareSumInvokedWhenRangeFilterKeepsNothing) {
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"A\"}, {\"X\";\"X\"}, {TRUE;\"text\"}, SUM, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 0.0);
}

TEST(PivotBy, BareAverageName) {
  // (A, X) = avg(10, 30) = 20; (A, Y) = 40; (B, X) = 20; (B, Y) = avg(60, 80) = 70.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"A\";\"B\";\"B\";\"B\"}, {\"X\";\"X\";\"Y\";\"X\";\"Y\";\"Y\"},"
      "         {10;30;40;20;60;80}, AVERAGE, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 20.0);  // (A, X)
  EXPECT_DOUBLE_EQ(Cell(v, 0, 2).as_number(), 40.0);  // (A, Y)
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 20.0);  // (B, X)
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 70.0);  // (B, Y)
}

TEST(PivotBy, BareCountAName) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, COUNTA, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 2).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 1.0);
}

TEST(PivotBy, UnknownBareNameYieldsNameError) {
  // `NOPE_NAME` resolves neither in NameEnv nor in the registry; the raw
  // NameRef evaluation surfaces #NAME?.
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, NOPE_NAME, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// ---------------------------------------------------------------------------
// Form B: name-bound LAMBDA via LET
// ---------------------------------------------------------------------------

TEST(PivotBy, NameBoundLambdaViaLet) {
  const Value v = EvalSrc(
      "=LET(agg, LAMBDA(v, SUM(v)),"
      "     PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "             {1;2;3;4}, agg, 0, 0, 0, 0, 0))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 4.0);
}

// ---------------------------------------------------------------------------
// field_headers ∈ {0, 1, 2, 3}
// ---------------------------------------------------------------------------

TEST(PivotBy, FieldHeadersZeroNoHeaders) {
  // No header row in inputs and no header row in output.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // First cell is a row key, not a header label.
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
}

TEST(PivotBy, FieldHeadersOneCopiesInputHeaders) {
  // Inputs have a header row; output also emits one. Top-left cell = the
  // row_fields header label ("Row").
  const Value v = EvalSrc(
      "=PIVOTBY({\"Row\";\"A\";\"A\";\"B\";\"B\"}, {\"Col\";\"X\";\"Y\";\"X\";\"Y\"},"
      "         {\"V\";1;2;3;4}, SUM, 1, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "Row");
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "Y");
}

TEST(PivotBy, FieldHeadersTwoSynthesizesDefaults) {
  // Inputs have no header row; output emits a synthesised "Field 1".
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 2, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "Field 1");
}

TEST(PivotBy, FieldHeadersThreeBothInputsHaveAndOutputEmits) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"R\";\"A\";\"B\"}, {\"C\";\"X\";\"Y\"},"
      "         {\"V\";1;2}, SUM, 3, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "R");
}

TEST(PivotBy, FieldHeadersOutOfRangeYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\"}, {\"X\"}, {1}, SUM, 5, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// row_total_depth
// ---------------------------------------------------------------------------

TEST(PivotBy, RowTotalDepthZeroNoColumnTotals) {
  // No "column totals" row anywhere.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // 2 body rows only.
  EXPECT_EQ(v.as_array_rows(), 2U);
}

TEST(PivotBy, RowTotalDepthPositiveOneColumnTotalsAtBottom) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 1, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  // Last row is the totals row: ["合計", X-total=4, Y-total=6].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 6.0);
}

TEST(PivotBy, RowTotalDepthNegativeOneColumnTotalsAtTop) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, -1, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 2).as_number(), 6.0);
}

TEST(PivotBy, RowTotalDepthOutOfRangeYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\"}, {\"X\"}, {1}, SUM, 0, 99, 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// col_total_depth
// ---------------------------------------------------------------------------

TEST(PivotBy, ColTotalDepthZeroNoRowTotals) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // No row-totals column.
  EXPECT_EQ(v.as_array_cols(), 3U);
}

TEST(PivotBy, ColTotalDepthPositiveOneRowTotalsOnRight) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0, 0, 1, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_cols(), 4U);
  // Last column on each row is the row-total: A->3, B->7.
  EXPECT_DOUBLE_EQ(Cell(v, 0, 3).as_number(), 3.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 3).as_number(), 7.0);
}

TEST(PivotBy, ColTotalDepthNegativeOneRowTotalsOnLeft) {
  // col_total_depth=-1 => row totals appear immediately to the right of
  // the row label column.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\";\"B\"}, {\"X\";\"Y\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0, 0, -1, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_cols(), 4U);
  // Col layout: [row_label, row_total, X, Y].
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 3.0);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 2).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 3).as_number(), 2.0);
}

TEST(PivotBy, ColTotalDepthOutOfRangeYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\"}, {\"X\"}, {1}, SUM, 0, 0, 0, 99, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// row_sort_order
// ---------------------------------------------------------------------------

TEST(PivotBy, RowSortOrderZeroPreservesFirstOccurrence) {
  // Row groups in input order: B, A, C. With sort_order=0 they stay that way.
  const Value v = EvalSrc(
      "=PIVOTBY({\"B\";\"A\";\"C\";\"A\"}, {\"X\";\"X\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "C");
}

TEST(PivotBy, RowSortOrderPositiveAscendingByRowTotal) {
  // Row totals: A=2+4=6, B=1, C=3. Ascending row-total: B, C, A.
  const Value v = EvalSrc(
      "=PIVOTBY({\"B\";\"A\";\"C\";\"A\"}, {\"X\";\"X\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0, 1, 1, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "C");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "A");
}

TEST(PivotBy, RowSortOrderNegativeDescendingByRowTotal) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"B\";\"A\";\"C\";\"A\"}, {\"X\";\"X\";\"X\";\"Y\"},"
      "         {1;2;3;4}, SUM, 0, 0, -1, 1, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "C");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
}

// ---------------------------------------------------------------------------
// col_sort_order
// ---------------------------------------------------------------------------

TEST(PivotBy, ColSortOrderZeroPreservesFirstOccurrence) {
  // Col groups in input order: Y, X, Z.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"A\";\"A\"}, {\"Y\";\"X\";\"Z\";\"X\"},"
      "         {1;2;3;4}, SUM, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Col layout: row_label | Y | X | Z.
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 1.0);  // Y
  EXPECT_DOUBLE_EQ(Cell(v, 0, 2).as_number(), 6.0);  // X = 2+4
  EXPECT_DOUBLE_EQ(Cell(v, 0, 3).as_number(), 3.0);  // Z
}

TEST(PivotBy, ColSortOrderPositiveAscendingByColTotal) {
  // Col totals: Y=1, X=6, Z=3. Ascending col-total: Y(1), Z(3), X(6).
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"A\";\"A\"}, {\"Y\";\"X\";\"Z\";\"X\"},"
      "         {1;2;3;4}, SUM, 0, 1, 0, 0, 1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Layout (with row_total_depth=1, col_total_depth=0): A row + bottom totals.
  // Header row first since field_headers default 3 + inputs without header
  // row triggers... actually field_headers here is 0 -> no header row, so
  // body row is row 0, and totals row would be at row 1. col_total_depth=0
  // means no row-totals column. Cols sorted ascending: row_label | Y | Z | X.
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 1.0);  // Y
  EXPECT_DOUBLE_EQ(Cell(v, 0, 2).as_number(), 3.0);  // Z
  EXPECT_DOUBLE_EQ(Cell(v, 0, 3).as_number(), 6.0);  // X
}

TEST(PivotBy, ColSortOrderNegativeDescendingByColTotal) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"A\";\"A\"}, {\"Y\";\"X\";\"Z\";\"X\"},"
      "         {1;2;3;4}, SUM, 0, 1, 0, 0, -1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Descending: X(6), Z(3), Y(1).
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 6.0);  // X
  EXPECT_DOUBLE_EQ(Cell(v, 0, 2).as_number(), 3.0);  // Z
  EXPECT_DOUBLE_EQ(Cell(v, 0, 3).as_number(), 1.0);  // Y
}

// ---------------------------------------------------------------------------
// filter_array
// ---------------------------------------------------------------------------

TEST(PivotBy, FilterArrayBasicIncludeExclude) {
  // Mask drops the second row (B, X, 3); only A rows and (B, Y, 4) survive.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\";\"A\";\"B\"}, {\"X\";\"X\";\"Y\";\"Y\"},"
      "         {1;3;2;4}, SUM, 0, 0, 0, 0, 0,"
      "         {TRUE;FALSE;TRUE;TRUE})");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  // (A, X) = 1, (A, Y) = 2, (B, Y) = 4. (B, X) row was masked out -> Blank.
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 2).as_number(), 2.0);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "B");
  // (B, X) has no surviving rows -> Blank cell.
  EXPECT_TRUE(Cell(v, 1, 1).is_blank()) << Cell(v, 1, 1).debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 4.0);
}

TEST(PivotBy, FilterArrayLengthMismatchYieldsValueError) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 0, 0, 0, 0, 0,"
      "         {TRUE;FALSE;TRUE})");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, FilterArrayAllExcludedYieldsValueError) {
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 0, 0, 0, 0, 0,"
      "         {FALSE;FALSE})");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Per-(row, col) error isolation
// ---------------------------------------------------------------------------

TEST(PivotBy, PerCellErrorIsolation) {
  // The aggregator divides 1 by the cell sum. (A, X) sums to 0 -> #DIV/0!
  // for that cell only; (A, Y) and (B, Y) sums are non-zero and produce
  // valid numbers.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\"}, {\"X\";\"Y\";\"Y\"},"
      "         {0;5;4}, LAMBDA(v, 1/SUM(v)), 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  // (A, X) errored.
  ASSERT_TRUE(Cell(v, 0, 1).is_error());
  EXPECT_EQ(Cell(v, 0, 1).as_error(), ErrorCode::Div0);
  // (A, Y) = 1/5 = 0.2.
  EXPECT_DOUBLE_EQ(Cell(v, 0, 2).as_number(), 0.2);
  // (B, X) is empty -> Blank, no aggregator call.
  EXPECT_TRUE(Cell(v, 1, 1).is_blank());
  // (B, Y) = 1/4 = 0.25.
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 0.25);
}

// ---------------------------------------------------------------------------
// ja-JP key folding (row + col)
// ---------------------------------------------------------------------------

TEST(PivotBy, JapaneseFoldingAppliesToRowAndColKeys) {
  // Half-width "ｱ" (U+FF71) and full-width "ア" (U+30A2) fold to the same
  // key. The pivot collapses both rows into a single (row_group, col_group)
  // cell totalling 30. Same applies to the col-key axis, where two
  // half-width katakana col-keys also fold together.
  const Value v = EvalSrc(
      "=PIVOTBY({\"\xef\xbd\xb1\";\"\xe3\x82\xa2\"},"
      "         {\"\xef\xbd\xb6\";\"\xe3\x82\xab\"},"  // ｶ folds to カ
      "         {10;20}, SUM, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 30.0);
}

// ---------------------------------------------------------------------------
// Row-count consistency
// ---------------------------------------------------------------------------

TEST(PivotBy, MismatchedRowCountsRowVsValuesYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\";\"C\"}, {\"X\";\"Y\";\"Z\"}, {1;2}, SUM, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, MismatchedRowCountsColVsValuesYieldsValueError) {
  // row_fields has 2 rows, values has 2 rows, col_fields has 3 -> #VALUE!.
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\";\"Z\"}, {1;2}, SUM, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Multi-column row_fields / col_fields / values
// ---------------------------------------------------------------------------
//
// The general K x L x V layout (K = row_fields->cols, L = col_fields->cols,
// V = values->cols, all >= 1) is documented in
// `src/eval/groupby_pivotby/pivotby.cpp` near the top of `eval_pivotby_lazy`.
// The fixtures here use Formulon's defaults (field_headers=3,
// row_total_depth=-1, col_total_depth=1) which differ from Mac Excel's
// defaults; that is a documented, intentional divergence preserving the
// pre-multi-col test contracts. When fixtures override the defaults
// (e.g. to fh=0), the expected output is computed from the spec layout.

TEST(PivotBy, MultiColumnRowFieldsBasic) {
  // K=2, L=1, V=1, fh=0. Formulon defaults for the rest:
  //   row_total_depth=1 (totals at BOTTOM), col_total_depth=1 (totals at RIGHT).
  // rows = L + 0 + nR + 1 = 1 + 0 + 2 + 1 = 4
  // cols = K + nC*V + V = 2 + 2 + 1 = 5
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\",\"P\";\"B\",\"Q\";\"A\",\"P\"}, {\"X\";\"Y\";\"X\"},"
      "         {1;2;3}, SUM, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cols(), 5U);
  // Row 0 (col-axis label row): [null, null, "X", "Y", "合計"].
  EXPECT_TRUE(Cell(v, 0, 0).is_blank());
  EXPECT_TRUE(Cell(v, 0, 1).is_blank());
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 4).as_text()), "合計");
  // Row 1 ((A,P) group): X=1+3=4, Y=blank, total=4.
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 1, 1).as_text()), "P");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 4.0);
  EXPECT_TRUE(Cell(v, 1, 3).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 1, 4).as_number(), 4.0);
  // Row 2 ((B,Q) group): X=blank, Y=2, total=2.
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "Q");
  EXPECT_TRUE(Cell(v, 2, 2).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 4).as_number(), 2.0);
  // Row 3 (bottom grand-total row): ["合計", null, 4, 2, 6].
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "合計");
  EXPECT_TRUE(Cell(v, 3, 1).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 4).as_number(), 6.0);
}

TEST(PivotBy, MultiColumnColFieldsBasic) {
  // K=1, L=2, V=1, fh=0.
  // rows = 2 + 0 + 2 + 1 = 5, cols = 1 + 2 + 1 = 4.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\";\"A\"}, {\"X\",\"M\";\"Y\",\"N\";\"X\",\"M\"},"
      "         {1;2;3}, SUM, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 5U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  // Row 0 (outermost col-axis level): [null, "X", "Y", "合計"].
  EXPECT_TRUE(Cell(v, 0, 0).is_blank());
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "合計");
  // Row 1 (innermost col-axis level): [null, "M", "N", null].
  EXPECT_TRUE(Cell(v, 1, 0).is_blank());
  EXPECT_EQ(std::string(Cell(v, 1, 1).as_text()), "M");
  EXPECT_EQ(std::string(Cell(v, 1, 2).as_text()), "N");
  EXPECT_TRUE(Cell(v, 1, 3).is_blank());
  // Row 2 (A): [A, 4, blank, 4].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 4.0);
  EXPECT_TRUE(Cell(v, 2, 2).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 4.0);
  // Row 3 (B): [B, blank, 2, 2].
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "B");
  EXPECT_TRUE(Cell(v, 3, 1).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 3).as_number(), 2.0);
  // Row 4 (bottom grand total): ["合計", 4, 2, 6].
  EXPECT_EQ(std::string(Cell(v, 4, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 4, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 2).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 3).as_number(), 6.0);
}

TEST(PivotBy, MultiColumnValuesBasic) {
  // K=1, L=1, V=2, fh=0.
  // rows = 1 + 0 + 2 + 1 = 4, cols = 1 + 2*2 + 2 = 7.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\";\"A\"}, {\"X\";\"Y\";\"X\"},"
      "         {1,10;2,20;3,30}, SUM, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cols(), 7U);
  // Row 0 (col-axis labels tiled V=2 per col group, "合計" tiled V=2):
  //   [null, "X", "X", "Y", "Y", "合計", "合計"].
  EXPECT_TRUE(Cell(v, 0, 0).is_blank());
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 4).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 5).as_text()), "合計");
  EXPECT_EQ(std::string(Cell(v, 0, 6).as_text()), "合計");
  // Row 1 (A): [A, 4, 40, blank, blank, blank, blank].
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 40.0);
  EXPECT_TRUE(Cell(v, 1, 3).is_blank());
  EXPECT_TRUE(Cell(v, 1, 4).is_blank());
  EXPECT_TRUE(Cell(v, 1, 5).is_blank());
  EXPECT_TRUE(Cell(v, 1, 6).is_blank());
  // Row 2 (B): [B, blank, blank, 2, 20, blank, blank].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
  EXPECT_TRUE(Cell(v, 2, 1).is_blank());
  EXPECT_TRUE(Cell(v, 2, 2).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 4).as_number(), 20.0);
  EXPECT_TRUE(Cell(v, 2, 5).is_blank());
  EXPECT_TRUE(Cell(v, 2, 6).is_blank());
  // Row 3 (bottom grand totals).
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 3, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 40.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 4).as_number(), 20.0);
  EXPECT_TRUE(Cell(v, 3, 5).is_blank());
  EXPECT_TRUE(Cell(v, 3, 6).is_blank());
}

TEST(PivotBy, MultiColumnAllAxes) {
  // K=2, L=2, V=2, fh=0.
  // rows = 2 + 0 + 2 + 1 = 5, cols = 2 + 2*2 + 2 = 8.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\",\"P\";\"B\",\"Q\";\"A\",\"P\"},"
      "         {\"X\",\"M\";\"Y\",\"N\";\"X\",\"M\"},"
      "         {1,10;2,20;3,30}, SUM, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 5U);
  EXPECT_EQ(v.as_array_cols(), 8U);
  // Row 0 (outer col-axis): [null, null, "X", "X", "Y", "Y",
  //                          "合計", "合計"].
  EXPECT_TRUE(Cell(v, 0, 0).is_blank());
  EXPECT_TRUE(Cell(v, 0, 1).is_blank());
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 4).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 5).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 0, 6).as_text()), "合計");
  EXPECT_EQ(std::string(Cell(v, 0, 7).as_text()), "合計");
  // Row 1 (inner col-axis): [null, null, "M", "M", "N", "N", null, null].
  EXPECT_TRUE(Cell(v, 1, 0).is_blank());
  EXPECT_TRUE(Cell(v, 1, 1).is_blank());
  EXPECT_EQ(std::string(Cell(v, 1, 2).as_text()), "M");
  EXPECT_EQ(std::string(Cell(v, 1, 3).as_text()), "M");
  EXPECT_EQ(std::string(Cell(v, 1, 4).as_text()), "N");
  EXPECT_EQ(std::string(Cell(v, 1, 5).as_text()), "N");
  EXPECT_TRUE(Cell(v, 1, 6).is_blank());
  EXPECT_TRUE(Cell(v, 1, 7).is_blank());
  // Row 2 ((A,P)): [A, P, 4, 40, null, null, null, null].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "P");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 40.0);
  EXPECT_TRUE(Cell(v, 2, 4).is_blank());
  EXPECT_TRUE(Cell(v, 2, 5).is_blank());
  EXPECT_TRUE(Cell(v, 2, 6).is_blank());
  EXPECT_TRUE(Cell(v, 2, 7).is_blank());
  // Row 3 ((B,Q)): [B, Q, null, null, 2, 20, null, null].
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 3, 1).as_text()), "Q");
  EXPECT_TRUE(Cell(v, 3, 2).is_blank());
  EXPECT_TRUE(Cell(v, 3, 3).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 3, 4).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 5).as_number(), 20.0);
  EXPECT_TRUE(Cell(v, 3, 6).is_blank());
  EXPECT_TRUE(Cell(v, 3, 7).is_blank());
  // Row 4 (bottom grand total): ["合計", null, 4, 40, 2, 20, null, null].
  EXPECT_EQ(std::string(Cell(v, 4, 0).as_text()), "合計");
  EXPECT_TRUE(Cell(v, 4, 1).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 4, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 3).as_number(), 40.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 4).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 5).as_number(), 20.0);
  EXPECT_TRUE(Cell(v, 4, 6).is_blank());
  EXPECT_TRUE(Cell(v, 4, 7).is_blank());
}

TEST(PivotBy, MultiColumnRowsWithFieldHeadersThree) {
  // K=2, L=1, V=1, fh=3 (inputs have header, output emits header).
  // rows = 1 col-field header + 1 col-axis + 1 output header + 1 total + 2 body = 6.
  const Value v = EvalSrc(
      "=PIVOTBY({\"R1\",\"R2\";\"A\",\"P\";\"B\",\"Q\";\"A\",\"P\"},"
      "         {\"C\";\"X\";\"Y\";\"X\"},"
      "         {\"V\";1;2;3}, SUM, 3)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 6U);
  EXPECT_EQ(v.as_array_cols(), 5U);
  // Row 0 (col-field header): [null, null, "C", null, null].
  EXPECT_TRUE(Cell(v, 0, 0).is_blank());
  EXPECT_TRUE(Cell(v, 0, 1).is_blank());
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "C");
  EXPECT_TRUE(Cell(v, 0, 3).is_blank());
  EXPECT_TRUE(Cell(v, 0, 4).is_blank());
  // Row 1 (col-axis labels from col_fields data row 0 reps):
  //   [null, null, "X", "Y", "合計"].
  EXPECT_TRUE(Cell(v, 1, 0).is_blank());
  EXPECT_TRUE(Cell(v, 1, 1).is_blank());
  EXPECT_EQ(std::string(Cell(v, 1, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 1, 3).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 1, 4).as_text()), "合計");
  // Row 2 (header row): ["R1", "R2", "V", "V", "V"].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "R1");
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "R2");
  EXPECT_EQ(std::string(Cell(v, 2, 2).as_text()), "V");
  EXPECT_EQ(std::string(Cell(v, 2, 3).as_text()), "V");
  EXPECT_EQ(std::string(Cell(v, 2, 4).as_text()), "V");
  // Row 3 ((A,P)): [A, P, 4, blank, 4].
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 3, 1).as_text()), "P");
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 4.0);
  EXPECT_TRUE(Cell(v, 3, 3).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 3, 4).as_number(), 4.0);
  // Row 4 ((B,Q)): [B, Q, blank, 2, 2].
  EXPECT_EQ(std::string(Cell(v, 4, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 4, 1).as_text()), "Q");
  EXPECT_TRUE(Cell(v, 4, 2).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 4, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 4).as_number(), 2.0);
  // Row 5 (bottom grand total).
  EXPECT_EQ(std::string(Cell(v, 5, 0).as_text()), "合計");
  EXPECT_TRUE(Cell(v, 5, 1).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 5, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 5, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 5, 4).as_number(), 6.0);
}

TEST(PivotBy, MultiColumnValuesWithFieldHeadersThree) {
  // K=1, L=1, V=2, fh=3.
  // rows = 1 col-field header + 1 col-axis + 1 output header + 1 total + 2 body = 6.
  const Value v = EvalSrc(
      "=PIVOTBY({\"Region\";\"A\";\"B\";\"A\"}, {\"Cat\";\"X\";\"Y\";\"X\"},"
      "         {\"S\",\"C\";1,10;2,20;3,30}, SUM, 3)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 6U);
  EXPECT_EQ(v.as_array_cols(), 7U);
  // Row 0: [null, "Cat", null, null, null, null, null].
  EXPECT_TRUE(Cell(v, 0, 0).is_blank());
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "Cat");
  EXPECT_TRUE(Cell(v, 0, 2).is_blank());
  EXPECT_TRUE(Cell(v, 0, 3).is_blank());
  EXPECT_TRUE(Cell(v, 0, 4).is_blank());
  EXPECT_TRUE(Cell(v, 0, 5).is_blank());
  EXPECT_TRUE(Cell(v, 0, 6).is_blank());
  // Row 1: [null, "X", "X", "Y", "Y", "合計", "合計"].
  EXPECT_TRUE(Cell(v, 1, 0).is_blank());
  EXPECT_EQ(std::string(Cell(v, 1, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 1, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 1, 3).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 1, 4).as_text()), "Y");
  EXPECT_EQ(std::string(Cell(v, 1, 5).as_text()), "合計");
  EXPECT_EQ(std::string(Cell(v, 1, 6).as_text()), "合計");
  // Row 2: ["Region", "S", "C", "S", "C", "S", "C"].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "Region");
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "S");
  EXPECT_EQ(std::string(Cell(v, 2, 2).as_text()), "C");
  EXPECT_EQ(std::string(Cell(v, 2, 3).as_text()), "S");
  EXPECT_EQ(std::string(Cell(v, 2, 4).as_text()), "C");
  EXPECT_EQ(std::string(Cell(v, 2, 5).as_text()), "S");
  EXPECT_EQ(std::string(Cell(v, 2, 6).as_text()), "C");
  // Row 3 (A): [A, 4, 40, blank, blank, blank, blank].
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 3, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 40.0);
  EXPECT_TRUE(Cell(v, 3, 3).is_blank());
  EXPECT_TRUE(Cell(v, 3, 4).is_blank());
  EXPECT_TRUE(Cell(v, 3, 5).is_blank());
  EXPECT_TRUE(Cell(v, 3, 6).is_blank());
  // Row 4 (B): [B, blank, blank, 2, 20, blank, blank].
  EXPECT_EQ(std::string(Cell(v, 4, 0).as_text()), "B");
  EXPECT_TRUE(Cell(v, 4, 1).is_blank());
  EXPECT_TRUE(Cell(v, 4, 2).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 4, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 4, 4).as_number(), 20.0);
  EXPECT_TRUE(Cell(v, 4, 5).is_blank());
  EXPECT_TRUE(Cell(v, 4, 6).is_blank());
  // Row 5 (bottom grand total).
  EXPECT_EQ(std::string(Cell(v, 5, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 5, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 5, 2).as_number(), 40.0);
  EXPECT_DOUBLE_EQ(Cell(v, 5, 3).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 5, 4).as_number(), 20.0);
  EXPECT_TRUE(Cell(v, 5, 5).is_blank());
  EXPECT_TRUE(Cell(v, 5, 6).is_blank());
}

TEST(PivotBy, MultiColumnValuesWithFieldHeadersTwoSynth) {
  // K=1, L=1, V=2, fh=2 (no input header, output emits synth labels).
  // rows = 1 + 1 + 2 + 1 = 5, cols = 7.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"B\";\"A\"}, {\"X\";\"Y\";\"X\"},"
      "         {1,10;2,20;3,30}, SUM, 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 5U);
  EXPECT_EQ(v.as_array_cols(), 7U);
  // Row 0 (col-axis labels from col_fields data row 0 reps):
  //   [null, "X", "X", "Y", "Y", "合計", "合計"].
  EXPECT_TRUE(Cell(v, 0, 0).is_blank());
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 5).as_text()), "合計");
  // Row 1 (synth header row):
  //   ["Field 1", "Value 1", "Value 2", "Value 1", "Value 2", "Value 1", "Value 2"].
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "Field 1");
  EXPECT_EQ(std::string(Cell(v, 1, 1).as_text()), "Value 1");
  EXPECT_EQ(std::string(Cell(v, 1, 2).as_text()), "Value 2");
  EXPECT_EQ(std::string(Cell(v, 1, 3).as_text()), "Value 1");
  EXPECT_EQ(std::string(Cell(v, 1, 4).as_text()), "Value 2");
  EXPECT_EQ(std::string(Cell(v, 1, 5).as_text()), "Value 1");
  EXPECT_EQ(std::string(Cell(v, 1, 6).as_text()), "Value 2");
  // Body sanity check (full coverage in MultiColumnValuesBasic).
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 4.0);
  EXPECT_TRUE(Cell(v, 2, 6).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 4, 1).as_number(), 4.0);
  EXPECT_TRUE(Cell(v, 4, 6).is_blank());
}

TEST(PivotBy, MultiColumnRowsWithFilterArray) {
  // K=2, L=1, V=1, fh=0, with a filter that drops the (B,Q) data row.
  // After filter: only (A,P) row group, only X col group. nR=1, nC=1.
  // rows = 1 + 0 + 1 + 1 = 3, cols = 2 + 1 + 1 = 4.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\",\"P\";\"B\",\"Q\";\"A\",\"P\"}, {\"X\";\"Y\";\"X\"},"
      "         {1;2;3}, SUM, 0, -1, 0, 1, 0, {TRUE;FALSE;TRUE})");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  // Row 0: [null, null, "X", "合計"].
  EXPECT_TRUE(Cell(v, 0, 0).is_blank());
  EXPECT_TRUE(Cell(v, 0, 1).is_blank());
  EXPECT_EQ(std::string(Cell(v, 0, 2).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 3).as_text()), "合計");
  // Row 1 (top grand total): ["合計", null, 4, 4].
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "合計");
  EXPECT_TRUE(Cell(v, 1, 1).is_blank());
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 3).as_number(), 4.0);
  // Row 2 ((A,P)): [A, P, 4, 4].
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "P");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 4.0);
}

// ---------------------------------------------------------------------------
// Bad arg domain checks
// ---------------------------------------------------------------------------

TEST(PivotBy, NonLambdaNonNameAggregatorYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\"}, {\"X\"}, {1}, 42, 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, LambdaWrongArityYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\"}, {\"X\"}, {1}, LAMBDA(a, b, a+b), 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, RowSortOrderAbsTwoYieldsValueError) {
  // First-commit scope: only ±1 / 0 are supported on the sort orders.
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 0, 0, 2, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, ColSortOrderAbsTwoYieldsValueError) {
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\";\"Y\"}, {1;2}, SUM, 0, 0, 0, 0, 2)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(PivotBy, TooFewArgsYieldsValueError) {
  // Fewer than 4 args is grammatically invalid (PIVOTBY requires the
  // first 4: row_fields, col_fields, values, function).
  const Value v = EvalSrc("=PIVOTBY({\"A\"}, {\"X\"}, {1})");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Aggregator returning multi-cell array -> #CALC! per cell
// ---------------------------------------------------------------------------

TEST(PivotBy, AggregatorReturningArrayYieldsCalcInThatCell) {
  // SEQUENCE returns multi-cell; each body cell becomes #CALC!.
  const Value v = EvalSrc(
      "=PIVOTBY({\"A\";\"A\";\"B\"}, {\"X\";\"Y\";\"X\"},"
      "         {1;2;3}, LAMBDA(v, SEQUENCE(1, 3)), 0, 0, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_TRUE(Cell(v, 0, 1).is_error());
  EXPECT_EQ(Cell(v, 0, 1).as_error(), ErrorCode::Calc);
  ASSERT_TRUE(Cell(v, 0, 2).is_error());
  EXPECT_EQ(Cell(v, 0, 2).as_error(), ErrorCode::Calc);
  // (B, X) is non-empty so the aggregator runs and surfaces #CALC! too.
  ASSERT_TRUE(Cell(v, 1, 1).is_error());
  EXPECT_EQ(Cell(v, 1, 1).as_error(), ErrorCode::Calc);
}

// ---------------------------------------------------------------------------
// Nested subtotals (|row_total_depth| == 2); column subtotals still fall back
// ---------------------------------------------------------------------------

TEST(PivotBy, RowTotalDepthTwoAddsASubtotalRowPerOuterRowGroup) {
  // Row keys are (outer, inner); the outer level is "A" (two rows) and "B"
  // (one row). Each outer group gets a subtotal row carrying its per-column
  // aggregates and its row total.
  const Value v =
      EvalSrc("=PIVOTBY({\"A\",\"x\";\"A\",\"y\";\"B\",\"x\"}, {\"X\";\"Y\";\"X\"}, {10;20;30}, SUM, 0, 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Col-axis label row + 3 body rows + 2 subtotals + grand total = 7.
  ASSERT_EQ(v.as_array_rows(), 7U);
  // 2 row-key cols + 2 col groups + grand-total col = 5.
  ASSERT_EQ(v.as_array_cols(), 5U);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 1, 1).as_text()), "x");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 10.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 3).as_number(), 20.0);
  // Subtotal for outer group A: 10 under X, 20 under Y, 30 in total.
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "A");
  EXPECT_TRUE(Cell(v, 3, 1).is_blank()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 10.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 3).as_number(), 20.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 4).as_number(), 30.0);
  // Outer group B has no Y data, so that cell stays blank.
  EXPECT_EQ(std::string(Cell(v, 5, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 5, 2).as_number(), 30.0);
  EXPECT_TRUE(Cell(v, 5, 3).is_blank()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 5, 4).as_number(), 30.0);
  // Grand total closes the block, promoted to 総計 now that subtotal rows
  // share the same column.
  EXPECT_EQ(std::string(Cell(v, 6, 0).as_text()), "総計");
  EXPECT_DOUBLE_EQ(Cell(v, 6, 4).as_number(), 60.0);
}

TEST(PivotBy, RowTotalDepthNegativeTwoPutsEverySubtotalAboveItsGroup) {
  const Value v =
      EvalSrc("=PIVOTBY({\"A\",\"x\";\"A\",\"y\";\"B\",\"x\"}, {\"X\";\"Y\";\"X\"}, {10;20;30}, SUM, 0, -2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 7U);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "総計");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 3, 1).as_text()), "x");
  EXPECT_EQ(std::string(Cell(v, 4, 1).as_text()), "y");
  EXPECT_EQ(std::string(Cell(v, 5, 0).as_text()), "B");
}

TEST(PivotBy, RowTotalDepthTwoWithOneRowKeyColumnKeepsTheGrandTotalOnlyLayout) {
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"A\";\"B\"}, {\"X\";\"Y\";\"X\"}, {10;20;30}, SUM, 0, 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // 2 body rows + grand total; the merged single-column layout emits no
  // label row when field_headers is 0.
  ASSERT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "合計");
}

TEST(PivotBy, RowTotalDepthTwoEmitsNoFallbackDiagnostic) {
  // Row subtotals are implemented, so the degraded-layout warning must not
  // fire on this path.
  testing::internal::CaptureStderr();
  const Value v =
      EvalSrc("=PIVOTBY({\"A\",\"x\";\"A\",\"y\";\"B\",\"x\"}, {\"X\";\"Y\";\"X\"}, {10;20;30}, SUM, 0, 2)");
  const std::string captured = testing::internal::GetCapturedStderr();
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(captured.find("subtotals_unsupported"), std::string::npos)
      << "unexpected fallback diagnostic; stderr was: " << captured;
}

TEST(PivotBy, ColTotalDepthTwoStillFallsBackAndSaysSo) {
  // Subtotal COLUMNS are not implemented; the ±2 column request degrades to
  // the grand-total-only layout and must announce it.
  testing::internal::CaptureStderr();
  const Value v = EvalSrc("=PIVOTBY({\"A\";\"B\"}, {\"X\",\"p\";\"X\",\"q\"}, {10;20}, SUM, 0, 1, 0, 2)");
  const std::string captured = testing::internal::GetCapturedStderr();
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_NE(captured.find("eval.pivotby.subtotals_unsupported"), std::string::npos)
      << "expected column-axis fallback diagnostic; stderr was: " << captured;
}

TEST(PivotBy, DefaultDepthEmitsNoSubtotalDiagnostic) {
  // The ordinary ±1 grand-total path must NOT emit the fallback warning.
  testing::internal::CaptureStderr();
  const Value v = EvalSrc("=PIVOTBY({\"A\",\"x\";\"A\",\"y\";\"B\",\"x\"}, {\"X\";\"Y\";\"X\"}, {10;20;30}, SUM, 0)");
  const std::string captured = testing::internal::GetCapturedStderr();
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(captured.find("eval.pivotby.subtotals_unsupported"), std::string::npos)
      << "unexpected fallback diagnostic on ±1 path; stderr was: " << captured;
}

}  // namespace
}  // namespace eval
}  // namespace formulon
