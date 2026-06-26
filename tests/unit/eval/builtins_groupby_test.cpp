// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tests for Excel 365's GROUPBY dynamic-array function. GROUPBY rides the
// lazy-dispatch table because its `function` argument has three accepted
// shapes (LAMBDA literal, name-bound LAMBDA via LET, bare function name)
// and at least the latter requires inspecting the raw AST before any
// evaluation can happen — `eval_node` on a bare `SUM` identifier would
// surface `#NAME?`.
//
// Coverage targets per the brief:
//
//   * Form A: inline LAMBDA literal aggregator.
//   * Form B: name-bound LAMBDA via LET.
//   * Form C: bare function name (`SUM`, `AVERAGE`, `COUNT`).
//   * Multi-column `values`: one input column -> one output column.
//   * `field_headers` ∈ {0, 1, 2, 3}: input/output header handling.
//   * `total_depth`: 0 (none), 1 (bottom), -1 (top default), 2 (subtotal
//     hint with multi-column row_fields).
//   * `sort_order`: 0 (preserve), positive (asc by aggregate), negative (desc).
//   * `filter_array`: include/exclude pattern.
//   * Per-group error isolation: one group's aggregator errors, others
//     remain valid.
//   * Argument domain checks: out-of-range `field_headers`, `total_depth`,
//     `sort_order`; mismatched row counts.
//   * Empty data after filter -> #CALC!.
//   * Group-key equality with ja-JP text (`fold_jp_text`).
//   * Aggregator returning array -> #CALC! for that cell.

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

TEST(GroupBy, SingleKeyLambdaAggregatorSum) {
  // Three groups (A=10+30=40, B=20, C=40) ordered by first occurrence.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\";\"C\"}, {10;20;30;40}, LAMBDA(v, SUM(v)), 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 2U);

  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 40.0);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 20.0);
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "C");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 40.0);
}

// ---------------------------------------------------------------------------
// Form C: bare function name aggregator
// ---------------------------------------------------------------------------

TEST(GroupBy, BareSumName) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {10;20;30}, SUM, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 40.0);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 20.0);
}

TEST(GroupBy, BareAverageName) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"A\";\"B\"}, {10;30;20}, AVERAGE, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 20.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 20.0);
}

TEST(GroupBy, BareCountAName) {
  // COUNTA counts non-Blank cells; here the input has no blanks so it
  // behaves like COUNT for the registered eager-dispatch path. (Bare COUNT
  // itself rides the lazy dispatch table and is intentionally not eligible
  // for Form C — the registry-only fallback ignores lazy entries.)
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\";\"A\"}, {1;2;3;4}, COUNTA, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 3.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 1.0);
}

TEST(GroupBy, UnknownBareNameYieldsNameError) {
  // A name that resolves neither in NameEnv nor in the function registry
  // surfaces #NAME? from the raw NameRef evaluation path.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {1;2}, NOPE_THIS_IS_NOT_A_FUNCTION, 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// ---------------------------------------------------------------------------
// Form B: name-bound LAMBDA via LET
// ---------------------------------------------------------------------------

TEST(GroupBy, NameBoundLambdaViaLet) {
  const Value v = EvalSrc(
      "=LET(agg, LAMBDA(v, SUM(v)),"
      "     GROUPBY({\"A\";\"B\";\"A\"}, {1;2;3}, agg, 0, 0, 0))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// Multi-column values: one aggregated output column per input column
// ---------------------------------------------------------------------------

TEST(GroupBy, MultiColumnValuesAggregateIndependently) {
  // values has two columns; the output adds two aggregated columns.
  // Group A: col1 (10+30)=40, col2 (1+3)=4. Group B: col1=20, col2=2.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {10,1;20,2;30,3}, SUM, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 40.0);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 2).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 20.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// field_headers ∈ {0, 1, 2, 3}
// ---------------------------------------------------------------------------

TEST(GroupBy, FieldHeadersZeroNoHeaders) {
  // No headers anywhere; only data.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {1;2;3}, SUM, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // First row should be a data group (key + agg), not a header.
  EXPECT_TRUE(Cell(v, 0, 0).is_text());
  EXPECT_TRUE(Cell(v, 0, 1).is_number());
}

TEST(GroupBy, FieldHeadersOneCopiesInputHeaders) {
  // Row 0 of inputs is the header row and is excluded from data.
  const Value v = EvalSrc("=GROUPBY({\"Grp\";\"A\";\"B\";\"A\"}, {\"Sales\";1;2;3}, SUM, 1, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Header row first.
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "Grp");
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "Sales");
  // Then per-group rows: A=4, B=2.
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 4.0);
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 2.0);
}

TEST(GroupBy, FieldHeadersTwoSynthesizesDefaults) {
  // Inputs have no headers; output emits "Field 1" / "Value 1" defaults.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {10;20}, SUM, 2, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "Field 1");
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "Value 1");
}

TEST(GroupBy, FieldHeadersThreeBothInputsHaveAndOutputEmits) {
  const Value v = EvalSrc("=GROUPBY({\"H1\";\"A\";\"B\"}, {\"V1\";10;20}, SUM, 3, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "H1");
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "V1");
}

TEST(GroupBy, FieldHeadersOutOfRangeYieldsValueError) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {1;2}, SUM, 5, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// total_depth
// ---------------------------------------------------------------------------

TEST(GroupBy, TotalDepthZeroNoTotals) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {10;20}, SUM, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Two group rows; no extra total row.
  EXPECT_EQ(v.as_array_rows(), 2U);
}

TEST(GroupBy, TotalDepthOneGrandTotalAtBottom) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {10;20}, SUM, 0, 1, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  // Last row is the grand total: "合計" / 30.
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 30.0);
}

TEST(GroupBy, TotalDepthNegativeOneGrandTotalAtTop) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {10;20}, SUM, 0, -1, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 30.0);
}

TEST(GroupBy, TotalDepthTwoFallsBackToGrandTotalOnlyLayout) {
  // Nested per-outer-group subtotals (|total_depth| == 2) are not yet
  // implemented; the impl falls back to the ±1 grand-total-only layout.
  // With all-distinct composite keys this yields three group rows plus a
  // single grand-total row at the bottom (no intermediate subtotal rows).
  const Value v = EvalSrc("=GROUPBY({\"X\",\"A\";\"X\",\"B\";\"Y\",\"A\"}, {10;20;30}, SUM, 0, 2, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 4U);  // 3 group rows + 1 grand total (no subtotals)
  ASSERT_EQ(v.as_array_cols(), 3U);  // 2 key cols + 1 value col
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 2).as_number(), 10.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 20.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 30.0);
  // Grand total at the bottom; "合計" is the ja-JP grand-total label.
  EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 60.0);
}

TEST(GroupBy, TotalDepthTwoEmitsNonSilentDiagnostic) {
  // The ±2 fallback must be observable, not silent: a structured-log
  // warning is emitted so callers can detect the degraded layout.
  testing::internal::CaptureStderr();
  const Value v = EvalSrc("=GROUPBY({\"X\",\"A\";\"X\",\"B\";\"Y\",\"A\"}, {10;20;30}, SUM, 0, 2, 0)");
  const std::string captured = testing::internal::GetCapturedStderr();
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_NE(captured.find("eval.groupby.subtotals_unsupported"), std::string::npos)
      << "expected ±2 fallback diagnostic; stderr was: " << captured;
}

TEST(GroupBy, TotalDepthOneEmitsNoSubtotalDiagnostic) {
  // The ordinary ±1 grand-total path must NOT emit the fallback warning.
  testing::internal::CaptureStderr();
  const Value v = EvalSrc("=GROUPBY({\"X\",\"A\";\"X\",\"B\";\"Y\",\"A\"}, {10;20;30}, SUM, 0, 1, 0)");
  const std::string captured = testing::internal::GetCapturedStderr();
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(captured.find("eval.groupby.subtotals_unsupported"), std::string::npos)
      << "unexpected fallback diagnostic on ±1 path; stderr was: " << captured;
}

TEST(GroupBy, TotalDepthOutOfRangeYieldsValueError) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {10;20}, SUM, 0, 99, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// sort_order
// ---------------------------------------------------------------------------

TEST(GroupBy, SortOrderZeroPreservesFirstOccurrence) {
  // Groups in input order: B, A, C.
  const Value v = EvalSrc("=GROUPBY({\"B\";\"A\";\"C\";\"A\"}, {1;2;3;4}, SUM, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "C");
}

TEST(GroupBy, SortOrderPositiveAscendingByAggregate) {
  // Aggregates: A=2+4=6, B=1, C=3. Ascending: B(1), C(3), A(6).
  const Value v = EvalSrc("=GROUPBY({\"B\";\"A\";\"C\";\"A\"}, {1;2;3;4}, SUM, 0, 0, 1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 1.0);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "C");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 3.0);
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 6.0);
}

TEST(GroupBy, SortOrderNegativeDescendingByAggregate) {
  // Same aggregates as above, descending: A(6), C(3), B(1).
  const Value v = EvalSrc("=GROUPBY({\"B\";\"A\";\"C\";\"A\"}, {1;2;3;4}, SUM, 0, 0, -1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 6.0);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "C");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
}

TEST(GroupBy, SortOrderOutOfRangeYieldsValueError) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {1;2}, SUM, 0, 0, 99)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// filter_array
// ---------------------------------------------------------------------------

TEST(GroupBy, FilterArrayBasicIncludeExclude) {
  // Mask drops the second row (B) so only A remains.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {10;20;30}, SUM, 0, 0, 0, {TRUE;FALSE;TRUE})");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 40.0);
}

TEST(GroupBy, FilterArrayLengthMismatchYieldsValueError) {
  // Mask has 4 cells, data has 3; mismatch -> #VALUE!.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {1;2;3}, SUM, 0, 0, 0, {TRUE;FALSE;TRUE;TRUE})");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(GroupBy, FilterArrayAllExcludedYieldsValueError) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {1;2}, SUM, 0, 0, 0, {FALSE;FALSE})");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Per-group error isolation
// ---------------------------------------------------------------------------

TEST(GroupBy, PerGroupErrorIsolation) {
  // The aggregator divides 1 by the group sum. Group A sums to 0 -> the
  // aggregator surfaces #DIV/0! for that group; group B sums to 5 -> the
  // aggregator returns 0.2. Mac Excel keeps both groups in the result; the
  // failing group shows the error verbatim and the rest are still computed.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {0;5}, LAMBDA(v, 1/SUM(v)), 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  // A's aggregate errored.
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  ASSERT_TRUE(Cell(v, 0, 1).is_error());
  EXPECT_EQ(Cell(v, 0, 1).as_error(), ErrorCode::Div0);
  // B's aggregate is fine.
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 0.2);
}

// ---------------------------------------------------------------------------
// Bad arg domain checks
// ---------------------------------------------------------------------------

TEST(GroupBy, MismatchedRowCountsYieldsValueError) {
  // row_fields has 3 rows, values has 2 -> #VALUE!.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"C\"}, {1;2}, SUM, 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(GroupBy, NonLambdaNonNameAggregatorYieldsValueError) {
  // A bare number as the aggregator slot is neither a Lambda value nor a
  // function name -> #VALUE!.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {1;2}, 42, 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(GroupBy, LambdaWrongArityYieldsValueError) {
  // GROUPBY requires a 1-arg lambda; a 2-arg lambda surfaces #VALUE!.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {1;2}, LAMBDA(a, b, a+b), 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Empty data
// ---------------------------------------------------------------------------

TEST(GroupBy, AllExcludedAfterFilterYieldsValueError) {
  const Value v = EvalSrc("=GROUPBY({\"A\"}, {10}, SUM, 0, 0, 0, {FALSE})");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// ja-JP key folding
// ---------------------------------------------------------------------------

TEST(GroupBy, JapaneseKeyFoldingHalfWidthMatchesFullWidth) {
  // Half-width "ｱ" (U+FF71) and full-width "ア" (U+30A2) should fold to the
  // same group key under fold_jp_text. The aggregator therefore sees both
  // rows in a single group totalling 30.
  const Value v = EvalSrc("=GROUPBY({\"\xef\xbd\xb1\";\"\xe3\x82\xa2\"}, {10;20}, SUM, 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 30.0);
}

// ---------------------------------------------------------------------------
// Multi-cell aggregator return -> #CALC! per cell
// ---------------------------------------------------------------------------

TEST(GroupBy, AggregatorReturningArrayYieldsCalcInThatCell) {
  // SEQUENCE returns a multi-cell array. GROUPBY has a single cell per
  // (group, value-column) slot, so each aggregated cell is #CALC!.
  // The remaining columns / rows are still rendered.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {1;2}, LAMBDA(v, SEQUENCE(1, 3)), 0, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_TRUE(Cell(v, 0, 1).is_error());
  EXPECT_EQ(Cell(v, 0, 1).as_error(), ErrorCode::Calc);
  ASSERT_TRUE(Cell(v, 1, 1).is_error());
  EXPECT_EQ(Cell(v, 1, 1).as_error(), ErrorCode::Calc);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
