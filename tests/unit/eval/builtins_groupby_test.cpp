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
//   * Group-key equality and order under the profile's text collation.
//   * Aggregator returning an array -> #VALUE!, returning a LAMBDA -> #CALC! for that cell.

#include <array>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "excel_profile.h"
#include "groupby_pivotby_excel_parity_helpers.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "test_eval_helpers.h"
#include "util/test_log_recorder.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

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

Value EvalSrcWithProfile(std::string_view src, ExcelProfile profile) {
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
  return evaluate(*root, eval_arena, default_registry(), test::context_with_profile(profile));
}

// Same as `EvalSrc`, evaluated against a caller-supplied workbook/sheet so
// a formula can reference real cells (e.g. a whole-column range).
Value EvalSrcIn(std::string_view src, const Workbook& wb, const Sheet& sheet) {
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
  EvalState state;
  return evaluate(*root, eval_arena, default_registry(), test::mac_context(wb, sheet, state));
}

bool IsPlaceholder(const Value& v) {
  return v.is_text() && v.as_text().empty();
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
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\";\"C\"}, {10;20;30;40}, LAMBDA(v, SUM(v)), 0, 0)");
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
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {10;20;30}, SUM, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 40.0);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 20.0);
}

TEST(GroupBy, BareAverageName) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"A\";\"B\"}, {10;30;20}, AVERAGE, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 20.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 20.0);
}

TEST(GroupBy, BareSumFiltersRangeSourcedNonNumbers) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"A\";\"B\";\"B\"}, {1;TRUE;\"text\";2}, SUM, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 2.0);
}

TEST(GroupBy, BareSumInvokedWhenRangeFilterKeepsNothing) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"A\"}, {TRUE;\"text\"}, SUM, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 0.0);
}

TEST(GroupBy, BareCountAName) {
  // COUNTA counts non-Blank cells; here the input has no blanks so it
  // behaves like COUNT for the registered eager-dispatch path. (Bare COUNT
  // itself rides the lazy dispatch table and is intentionally not eligible
  // for Form C — the registry-only fallback ignores lazy entries.)
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\";\"A\"}, {1;2;3;4}, COUNTA, 0, 0)");
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
      "     GROUPBY({\"A\";\"B\";\"A\"}, {1;2;3}, agg, 0, 0))");
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
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {10,1;20,2;30,3}, SUM, 0, 0)");
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
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {1;2;3}, SUM, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // First row should be a data group (key + agg), not a header.
  EXPECT_TRUE(Cell(v, 0, 0).is_text());
  EXPECT_TRUE(Cell(v, 0, 1).is_number());
}

TEST(GroupBy, FieldHeadersOneConsumesInputHeadersWithoutShowingThem) {
  // Row 0 of inputs is the header row; it is excluded from data and not shown.
  const Value v = EvalSrc("=GROUPBY({\"Grp\";\"A\";\"B\";\"A\"}, {\"Sales\";1;2;3}, SUM, 1, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 4.0);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 2.0);
}

TEST(GroupBy, FieldHeadersOmittedDetectsATextHeaderRow) {
  const Value with_header = EvalSrc("=GROUPBY({\"k\";\"b\";\"a\";\"b\"}, {\"v\";1;2;3}, SUM)");
  ASSERT_TRUE(with_header.is_array()) << with_header.debug_to_string();
  EXPECT_EQ(with_header.as_array_rows(), 3U);
  EXPECT_EQ(std::string(Cell(with_header, 0, 0).as_text()), "a");
  EXPECT_DOUBLE_EQ(Cell(with_header, 0, 1).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(with_header, 1, 1).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(Cell(with_header, 2, 1).as_number(), 6.0);
  const Value without = EvalSrc("=GROUPBY({\"b\";\"a\";\"b\"}, {1;2;3}, SUM)");
  ASSERT_TRUE(without.is_array()) << without.debug_to_string();
  EXPECT_EQ(without.as_array_rows(), 3U);
  EXPECT_EQ(std::string(Cell(without, 0, 0).as_text()), "a");
}

TEST(GroupBy, SortOrderOmittedSortsKeysAscending) {
  Value v = EvalSrc("=GROUPBY({\"b\";\"a\";\"c\";\"a\"}, {1;2;3;4}, SUM, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "a");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 6.0);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "b");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "c");
  v = EvalSrc("=GROUPBY({3;1;2}, {1;2;3}, SUM, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 0).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 0).as_number(), 3.0);
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 1.0);
}

TEST(GroupBy, FieldHeadersTwoSynthesizesDefaults) {
  // Inputs have no headers; output emits generated labels.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {10;20}, SUM, 2, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "行フィールド 1");
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "値 1");
}

TEST(GroupBy, FieldHeadersThreeBothInputsHaveAndOutputEmits) {
  const Value v = EvalSrc("=GROUPBY({\"H1\";\"A\";\"B\"}, {\"V1\";10;20}, SUM, 3, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "H1");
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "V1");
}

TEST(GroupByProfile, GeneratedLabelsFollowWorkbookProfile) {
  struct ProfileCase {
    ExcelProfile profile;
    const char* id;
    const char* row_field;
    const char* value;
    const char* grand_total;
  };
  constexpr std::array<ProfileCase, 4> profiles = {{
      {mac_365_ja_jp_profile(), "mac-365-ja_JP", "行フィールド 1", "値 1", "合計"},
      {win_365_ja_jp_profile(), "win-365-ja_JP", "行フィールド 1", "値 1", "合計"},
      {mac_365_en_us_profile(), "mac-365-en_US", "Row Field 1", "Value 1", "Total"},
      {win_365_en_us_profile(), "win-365-en_US", "Row Field 1", "Value 1", "Total"},
  }};
  for (const ProfileCase& test_case : profiles) {
    SCOPED_TRACE(test_case.id);
    const Value v = EvalSrcWithProfile("=GROUPBY({\"A\";\"B\"},{1;2},SUM,2,1)", test_case.profile);
    ASSERT_TRUE(v.is_array()) << v.debug_to_string();
    ASSERT_EQ(v.as_array_rows(), 4U);
    ASSERT_EQ(v.as_array_cols(), 2U);
    EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), test_case.row_field);
    EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), test_case.value);
    EXPECT_EQ(std::string(Cell(v, 3, 0).as_text()), test_case.grand_total);
  }
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
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {10;20}, SUM, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // Two group rows; no extra total row.
  EXPECT_EQ(v.as_array_rows(), 2U);
}

TEST(GroupBy, TotalDepthOneGrandTotalAtBottom) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {10;20}, SUM, 0, 1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  // Last row is the grand total: "合計" / 30.
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 30.0);
}

TEST(GroupBy, TotalDepthDefaultsToGrandTotalAtBottom) {
  const Value v = EvalSrc("=GROUPBY({\"a\";\"a\";\"b\"}, {1;2;3}, SUM)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 3.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 3.0);
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 6.0);
}

TEST(GroupBy, TotalDepthNegativeOneGrandTotalAtTop) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {10;20}, SUM, 0, -1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "合計");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 30.0);
}

TEST(GroupBy, TotalDepthTwoAddsASubtotalRowPerOuterGroup) {
  // The outer level is the first key column, so X covers two groups and Y
  // one. Each outer group's rows are followed by its subtotal, and the
  // grand total closes the block.
  const Value v = EvalSrc("=GROUPBY({\"X\",\"A\";\"X\",\"B\";\"Y\",\"A\"}, {10;20;30}, SUM, 0, 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 6U);  // 3 group rows + 2 subtotals + 1 grand total
  ASSERT_EQ(v.as_array_cols(), 3U);  // 2 key cols + 1 value col
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "X");
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 2).as_number(), 10.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 20.0);
  // Subtotal for outer group X. The row restates the outer key verbatim and
  // blanks the inner key column, as on the grand-total row.
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "X");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 1))) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 2, 2).as_number(), 30.0);
  EXPECT_DOUBLE_EQ(Cell(v, 3, 2).as_number(), 30.0);
  EXPECT_EQ(std::string(Cell(v, 4, 0).as_text()), "Y");
  EXPECT_DOUBLE_EQ(Cell(v, 4, 2).as_number(), 30.0);
  // With subtotals in the same column the ja-JP grand total moves up to
  // 総計 so the two levels stay distinguishable.
  EXPECT_EQ(std::string(Cell(v, 5, 0).as_text()), "総計");
  EXPECT_DOUBLE_EQ(Cell(v, 5, 2).as_number(), 60.0);
}

TEST(GroupBy, TotalDepthNegativeTwoPutsEverySubtotalAboveItsGroup) {
  const Value v = EvalSrc("=GROUPBY({\"X\",\"A\";\"X\",\"B\";\"Y\",\"A\"}, {10;20;30}, SUM, 0, -2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 6U);
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "総計");
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "X");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 30.0);
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 3, 1).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 4, 0).as_text()), "Y");
}

TEST(GroupBy, TotalDepthBeyondTheKeyColumnsIsValueError) {
  // |depth| >= 2 needs that many key columns; fractional depths are rejected.
  const char* const formulas[] = {
      "=GROUPBY({\"A\";\"B\"}, {10;20}, SUM, 0, 2)",
      "=GROUPBY({\"A\";\"B\"}, {10;20}, SUM, 0, -2)",
      "=GROUPBY({\"A\",\"a\";\"B\",\"b\"}, {10;20}, SUM, 0, 3)",
      "=GROUPBY({\"A\",\"a\";\"B\",\"b\"}, {10;20}, SUM, 0, -3)",
      "=GROUPBY({\"A\",\"a\";\"B\",\"b\"}, {10;20}, SUM, 0, 2.5)",
      "=GROUPBY({\"A\",\"a\";\"B\",\"b\"}, {10;20}, SUM, 0, -2.5)",
  };
  for (const char* formula : formulas) {
    const Value v = EvalSrc(formula);
    ASSERT_TRUE(v.is_error()) << formula << ": " << v.debug_to_string();
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << formula;
  }
}

TEST(GroupBy, TotalDepthThreeNestsPrefixSubtotalsAfterTheirBlocks) {
  const Value v = EvalSrc(
      "=GROUPBY({\"x\",\"p\",\"m\";\"x\",\"p\",\"n\";\"x\",\"q\",\"n\";\"y\",\"p\",\"m\"}, {1;5;2;7}, SUM, 0, 3)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 10U);
  ASSERT_EQ(v.as_array_cols(), 4U);
  const char* const keys[10][3] = {{"x", "p", "m"}, {"x", "p", "n"}, {"x", "p", ""}, {"x", "q", "n"}, {"x", "q", ""},
                                   {"x", "", ""},   {"y", "p", "m"}, {"y", "p", ""}, {"y", "", ""},   {"総計", "", ""}};
  const double sums[10] = {1, 5, 6, 2, 2, 8, 7, 7, 7, 15};
  for (std::uint32_t r = 0; r < 10U; ++r) {
    for (std::uint32_t c = 0; c < 3U; ++c) {
      EXPECT_EQ(std::string(Cell(v, r, c).as_text()), keys[r][c]) << "row " << r << " col " << c;
    }
    EXPECT_DOUBLE_EQ(Cell(v, r, 3).as_number(), sums[r]) << "row " << r;
  }
}

TEST(GroupBy, TotalDepthNegativeThreeNestsPrefixSubtotalsBeforeTheirBlocks) {
  const Value v = EvalSrc(
      "=GROUPBY({\"x\",\"p\",\"m\";\"x\",\"p\",\"n\";\"x\",\"q\",\"n\";\"y\",\"p\",\"m\"}, {1;5;2;7}, SUM, 0, -3)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 10U);
  const char* const keys[10][3] = {{"総計", "", ""}, {"x", "", ""},   {"x", "p", ""}, {"x", "p", "m"}, {"x", "p", "n"},
                                   {"x", "q", ""},   {"x", "q", "n"}, {"y", "", ""},  {"y", "p", ""},  {"y", "p", "m"}};
  const double sums[10] = {15, 8, 6, 1, 5, 2, 2, 7, 7, 7};
  for (std::uint32_t r = 0; r < 10U; ++r) {
    for (std::uint32_t c = 0; c < 3U; ++c) {
      EXPECT_EQ(std::string(Cell(v, r, c).as_text()), keys[r][c]) << "row " << r << " col " << c;
    }
    EXPECT_DOUBLE_EQ(Cell(v, r, 3).as_number(), sums[r]) << "row " << r;
  }
}

TEST(GroupBy, TotalDepthThreeWithFourKeysEmitsNoPrefixThreeRows) {
  const Value v = EvalSrc(
      "=GROUPBY({\"x\",\"p\",\"m\",\"a\";\"x\",\"p\",\"m\",\"b\";\"x\",\"q\",\"n\",\"a\"}, {1;2;4}, SUM, 0, 3)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  // 3 groups + prefix-1 and prefix-2 subtotals for x, prefix-2 for (x,q) + grand total.
  ASSERT_EQ(v.as_array_rows(), 7U);
  ASSERT_EQ(v.as_array_cols(), 5U);
  // Rows 0-1 are the (x,p,m) leaves; the next row rolls up (x,p), not (x,p,m).
  EXPECT_EQ(std::string(Cell(v, 1, 3).as_text()), "b");
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "p");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 2, 2))) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 2, 4).as_number(), 3.0);
  EXPECT_EQ(std::string(Cell(v, 4, 1).as_text()), "q");
  EXPECT_TRUE(IsPlaceholder(Cell(v, 4, 2))) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 4, 4).as_number(), 4.0);
  EXPECT_TRUE(IsPlaceholder(Cell(v, 5, 1))) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 5, 4).as_number(), 7.0);
  EXPECT_EQ(std::string(Cell(v, 6, 0).as_text()), "総計");
}

TEST(GroupBy, SubtotalsKeepEachOuterGroupContiguousUnderASort) {
  // The hierarchy is primary: a value sort ranks the outer groups by their
  // own totals and orders the members within each, never interleaving them.
  const Value v = EvalSrc("=GROUPBY({\"X\",\"A\";\"Y\",\"B\";\"X\",\"C\"}, {10;5;1}, SUM, 0, 2, 3)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 6U);
  // Ascending: outer group Y (5) leads X (11); X's members run C (1), A (10).
  EXPECT_EQ(std::string(Cell(v, 0, 1).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "Y");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 2).as_number(), 5.0);
  EXPECT_EQ(std::string(Cell(v, 2, 1).as_text()), "C");
  EXPECT_EQ(std::string(Cell(v, 3, 1).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 4, 0).as_text()), "X");
  EXPECT_DOUBLE_EQ(Cell(v, 4, 2).as_number(), 11.0);
}

TEST(GroupBy, TotalDepthTwoEmitsNoDiagnostic) {
  // The ±2 path is implemented, so it must not warn about a degraded layout
  // the way the earlier grand-total-only fallback did. The assertion is
  // against a log sink rather than captured stderr: logging ships off, so a
  // stderr assertion would hold even if the code did warn.
  ::formulon::test::LogRecorder log;
  ASSERT_TRUE(log.probe_and_clear()) << "log sink is not carrying records; the assertion below would be vacuous";
  const Value v = EvalSrc("=GROUPBY({\"X\",\"A\";\"X\",\"B\";\"Y\",\"A\"}, {10;20;30}, SUM, 0, 2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_TRUE(log.empty()) << "unexpected diagnostic: " << log.joined();
}

TEST(GroupBy, TotalDepthOutOfRangeYieldsValueError) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {10;20}, SUM, 0, 99, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// sort_order
// ---------------------------------------------------------------------------

TEST(GroupBy, OmittedSortOrderSortsKeysAscending) {
  // Input order B, A, C comes out A, B, C.
  const Value v = EvalSrc("=GROUPBY({\"B\";\"A\";\"C\";\"A\"}, {1;2;3;4}, SUM, 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "C");
}

// A slot left empty between commas is an omission, not the value 0, so it
// reaches the same default rather than the rejection below.
TEST(GroupBy, EmptySortOrderSlotIsAnOmission) {
  const Value v = EvalSrc("=GROUPBY({\"B\";\"A\";\"C\";\"A\"}, {1;2;3;4}, SUM, 0, 0, , {TRUE;TRUE;TRUE;TRUE})");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "C");
}

// The argument is a signed column index and that domain has no zero member,
// so spelling the default out loud is rejected even though omitting it is
// not. Matches Excel (groupby_sort_order_zero_rejected).
TEST(GroupBy, SuppliedSortOrderZeroYieldsValueError) {
  const Value v = EvalSrc("=GROUPBY({\"B\";\"A\";\"C\";\"A\"}, {1;2;3;4}, SUM, 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// A fraction truncates toward zero before the domain check, so it lands on
// the same rejection rather than resolving to a column.
TEST(GroupBy, SortOrderTruncatingToZeroYieldsValueError) {
  const Value v = EvalSrc("=GROUPBY({\"B\";\"A\";\"C\";\"A\"}, {1;2;3;4}, SUM, 0, 0, 0.5)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// An unresolvable aggregator wins over the sort_order check. Excel agrees
// (groupby_unknown_name_error carries sort_order 0 and still answers
// #NAME?), so the two checks are ordered, not interchangeable.
TEST(GroupBy, UnknownAggregatorOutranksASuppliedSortOrderZero) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {1;2}, NOPE_THIS_IS_NOT_A_FUNCTION, 0, 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// sort_order numbers the output columns with the key columns first, so with
// one key column 2 names the aggregate.
TEST(GroupBy, SortOrderPositiveAscendingByAggregate) {
  // Aggregates: A=2+4=6, B=1, C=3. Ascending: B(1), C(3), A(6).
  const Value v = EvalSrc("=GROUPBY({\"B\";\"A\";\"C\";\"A\"}, {1;2;3;4}, SUM, 0, 0, 2)");
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
  const Value v = EvalSrc("=GROUPBY({\"B\";\"A\";\"C\";\"A\"}, {1;2;3;4}, SUM, 0, 0, -2)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 6.0);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "C");
  EXPECT_EQ(std::string(Cell(v, 2, 0).as_text()), "B");
}

TEST(GroupBy, SortOrderKeyColumnSetsTheKeyDirection) {
  const Value asc = EvalSrc("=GROUPBY({\"B\";\"A\";\"C\";\"A\"}, {1;2;3;4}, SUM, 0, 0, 1)");
  ASSERT_TRUE(asc.is_array()) << asc.debug_to_string();
  EXPECT_EQ(std::string(Cell(asc, 0, 0).as_text()), "A");
  EXPECT_EQ(std::string(Cell(asc, 2, 0).as_text()), "C");
  const Value desc = EvalSrc("=GROUPBY({\"B\";\"A\";\"C\";\"A\"}, {1;2;3;4}, SUM, 0, 0, -1)");
  ASSERT_TRUE(desc.is_array()) << desc.debug_to_string();
  EXPECT_EQ(std::string(Cell(desc, 0, 0).as_text()), "C");
  EXPECT_EQ(std::string(Cell(desc, 1, 0).as_text()), "B");
  EXPECT_EQ(std::string(Cell(desc, 2, 0).as_text()), "A");
}

// A descending inner key keeps the outer key outermost and ascending:
// x/b, x/a, y/b, y/a.
TEST(GroupBy, SortOrderVectorDescendingInnerKeyKeepsTheOuterKeyFirst) {
  const Value v = EvalSrc("=GROUPBY({\"x\",\"b\";\"y\",\"a\";\"x\",\"a\";\"y\",\"b\"}, {1;2;3;4}, SUM, 0, 0, {-2})");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 4U);
  const char* want[4][2] = {{"x", "b"}, {"x", "a"}, {"y", "b"}, {"y", "a"}};
  const double sums[4] = {1.0, 3.0, 4.0, 2.0};
  for (std::uint32_t r = 0; r < 4U; ++r) {
    EXPECT_EQ(std::string(Cell(v, r, 0).as_text()), want[r][0]) << r;
    EXPECT_EQ(std::string(Cell(v, r, 1).as_text()), want[r][1]) << r;
    EXPECT_DOUBLE_EQ(Cell(v, r, 2).as_number(), sums[r]) << r;
  }
}

TEST(GroupBy, SortOrderVectorNamingAValueColumnIsValueError) {
  const Value v = EvalSrc("=GROUPBY({\"x\",\"b\";\"y\",\"a\";\"x\",\"a\";\"y\",\"b\"}, {1;2;3;4}, SUM, 0, 0, {1,-3})");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(GroupBy, SortOrderPastTheLastColumnIsValueError) {
  const Value v = EvalSrc("=GROUPBY({\"B\";\"A\";\"C\";\"A\"}, {1;2;3;4}, SUM, 0, 0, 3)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
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
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {10;20;30}, SUM, 0, 0, , {TRUE;FALSE;TRUE})");
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

TEST(GroupBy, FilterArrayWholeColumnRangePopulatedInOneRowIsValueError) {
  // C:E is declared 2D (every row x 3 columns) even though only row 1 is
  // populated; the walked array (1 populated row x 3 cols) must not be
  // misread as a 1D vector just because the populated extent happens to
  // collapse to one row.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 2, Value::boolean(true));
  sheet.set_cell_value(0, 3, Value::boolean(true));
  sheet.set_cell_value(0, 4, Value::boolean(true));
  const Value v = EvalSrcIn("=GROUPBY({\"A\";\"B\";\"A\"}, {10;20;30}, SUM, 0, 0, , C:E)", wb, sheet);
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
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {0;5}, LAMBDA(v, 1/SUM(v)), 0, 0)");
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
  // GROUPBY calls the aggregator with one or two arguments; a lambda that
  // requires three cannot be called and surfaces #VALUE!.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {1;2}, LAMBDA(a, b, c, a+b+c), 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// A two-parameter aggregator receives (group values, all filtered values of
// the column) in every cell: groups, subtotals and the grand total.
TEST(GroupBy, PercentofReceivesTheGroupAndTheWholeColumn) {
  for (const char* agg : {"PERCENTOF", "LAMBDA(x,y,PERCENTOF(x,y))"}) {
    const Value v = EvalSrc(std::string("=GROUPBY({\"a\";\"b\";\"a\"},{1;2;3},") + agg + ",0,1)");
    ASSERT_TRUE(v.is_array()) << agg << " -> " << v.debug_to_string();
    ASSERT_EQ(v.as_array_rows(), 3U) << agg;
    EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 4.0 / 6.0) << agg;
    EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 2.0 / 6.0) << agg;
    EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 1.0) << agg;
  }
}

// Every aggregator requiring exactly two arguments, built-ins included, is
// called as (group values, whole value column).
TEST(GroupBy, TwoArgumentAggregatorsReceiveTheWholeColumn) {
  const std::string keys = "{\"a\";\"a\";\"b\";\"b\";\"c\";\"c\"},{1;2;3;4;5;6},";
  const Value rows = EvalSrc("=GROUPBY(" + keys + "LAMBDA(x,y,ROWS(x)&\"/\"&ROWS(y)))");
  ASSERT_TRUE(rows.is_array()) << rows.debug_to_string();
  ASSERT_EQ(rows.as_array_rows(), 4U);
  for (std::uint32_t r = 0; r < 3U; ++r) {
    EXPECT_EQ(std::string(Cell(rows, r, 1).as_text()), "2/6") << r;
  }
  EXPECT_EQ(std::string(Cell(rows, 3, 1).as_text()), "6/6");

  const Value correl = EvalSrc("=GROUPBY(" + keys + "CORREL)");
  ASSERT_TRUE(correl.is_array()) << correl.debug_to_string();
  for (std::uint32_t r = 0; r < 3U; ++r) {
    ASSERT_TRUE(Cell(correl, r, 1).is_error()) << r;
    EXPECT_EQ(Cell(correl, r, 1).as_error(), ErrorCode::NA) << r;
  }
  EXPECT_DOUBLE_EQ(Cell(correl, 3, 1).as_number(), 1.0);

  const Value covariance = EvalSrc("=GROUPBY(" + keys + "COVARIANCE.P)");
  ASSERT_TRUE(covariance.is_array()) << covariance.debug_to_string();
  EXPECT_EQ(Cell(covariance, 0, 1).as_error(), ErrorCode::NA);
  EXPECT_DOUBLE_EQ(Cell(covariance, 3, 1).as_number(), 2.9166666666666665);

  const Value percent = EvalSrc("=GROUPBY(" + keys + "PERCENTOF)");
  ASSERT_TRUE(percent.is_array()) << percent.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(percent, 0, 1).as_number(), 3.0 / 21.0);
  EXPECT_DOUBLE_EQ(Cell(percent, 1, 1).as_number(), 7.0 / 21.0);
  EXPECT_DOUBLE_EQ(Cell(percent, 2, 1).as_number(), 11.0 / 21.0);
  EXPECT_DOUBLE_EQ(Cell(percent, 3, 1).as_number(), 1.0);
}

// A built-in whose minimum is two arguments is called with two even when it
// takes more (XIRR, FORECAST.ETS.SEASONALITY); IF is called with one.
TEST(GroupBy, BuiltinWithMinimumTwoArgumentsReceivesTheWholeColumn) {
  const std::string keys = "{\"a\";\"a\";\"b\";\"b\"},{1;2;3;4},";
  const Value xirr = EvalSrc("=GROUPBY(" + keys + "XIRR)");
  ASSERT_TRUE(xirr.is_array()) << xirr.debug_to_string();
  ASSERT_EQ(xirr.as_array_rows(), 3U);
  for (std::uint32_t r = 0; r < 3U; ++r) {
    ASSERT_TRUE(Cell(xirr, r, 1).is_error()) << r;
    EXPECT_EQ(Cell(xirr, r, 1).as_error(), ErrorCode::Num) << r;
  }

  const Value seasonality = EvalSrc("=GROUPBY(" + keys + "FORECAST.ETS.SEASONALITY)");
  ASSERT_TRUE(seasonality.is_array()) << seasonality.debug_to_string();
  ASSERT_EQ(seasonality.as_array_rows(), 3U);
  EXPECT_EQ(Cell(seasonality, 0, 1).as_error(), ErrorCode::NA);
  EXPECT_EQ(Cell(seasonality, 1, 1).as_error(), ErrorCode::NA);
  EXPECT_DOUBLE_EQ(Cell(seasonality, 2, 1).as_number(), 0.0);

  const Value if_rows = EvalSrc("=GROUPBY(" + keys + "IF)");
  ASSERT_TRUE(if_rows.is_array()) << if_rows.debug_to_string();
  ASSERT_EQ(if_rows.as_array_rows(), 3U);
  for (std::uint32_t r = 0; r < 3U; ++r) {
    ASSERT_TRUE(Cell(if_rows, r, 1).is_error()) << r;
    EXPECT_EQ(Cell(if_rows, r, 1).as_error(), ErrorCode::Value) << r;
  }
}

TEST(GroupBy, PercentofSubtotalsAndFilterUseTheFilteredColumn) {
  const Value sub = EvalSrc("=GROUPBY({\"x\",\"b\";\"y\",\"a\";\"x\",\"a\";\"y\",\"b\"},{1;2;3;4},PERCENTOF,0,2)");
  ASSERT_TRUE(sub.is_array()) << sub.debug_to_string();
  ASSERT_EQ(sub.as_array_rows(), 7U);
  const double want[7] = {0.3, 0.1, 0.4, 0.2, 0.4, 0.6, 1.0};
  for (std::uint32_t r = 0; r < 7U; ++r) {
    EXPECT_DOUBLE_EQ(Cell(sub, r, 2).as_number(), want[r]) << r;
  }
  const Value filtered = EvalSrc("=GROUPBY({\"a\";\"b\";\"a\";\"c\"},{1;2;3;4},PERCENTOF,0,1,,{TRUE;TRUE;TRUE;FALSE})");
  ASSERT_TRUE(filtered.is_array()) << filtered.debug_to_string();
  ASSERT_EQ(filtered.as_array_rows(), 3U);
  EXPECT_DOUBLE_EQ(Cell(filtered, 0, 1).as_number(), 4.0 / 6.0);
  EXPECT_DOUBLE_EQ(Cell(filtered, 2, 1).as_number(), 1.0);
}

// ---------------------------------------------------------------------------
// Optional (`[name]`) params on the aggregator lambda
//
// A lambda requiring one parameter gets one argument per group, so it is
// callable when `param_count - optional_count <= 1 <= param_count`. The
// params it does not supply bind to the sentinel `ISOMITTED` detects.
// ---------------------------------------------------------------------------

TEST(GroupBy, LambdaWithTrailingOptionalIsAccepted) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {10;20;30}, LAMBDA(v, [u], SUM(v)), 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_EQ(std::string(Cell(v, 0, 0).as_text()), "A");
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 40.0);
  EXPECT_EQ(std::string(Cell(v, 1, 0).as_text()), "B");
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 20.0);
}

TEST(GroupBy, OptionalParamLambdaMatchesPlainLambda) {
  // `EvalSrc` resets the shared arenas per call, so the first result is
  // copied out as plain doubles before the second formula runs.
  double plain[2] = {0.0, 0.0};
  {
    const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {10;20;30}, LAMBDA(v, SUM(v)), 0, 0)");
    ASSERT_TRUE(v.is_array()) << v.debug_to_string();
    ASSERT_EQ(v.as_array_rows(), 2U);
    plain[0] = Cell(v, 0, 1).as_number();
    plain[1] = Cell(v, 1, 1).as_number();
  }
  const Value with_opt = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {10;20;30}, LAMBDA(v, [u], SUM(v)), 0, 0)");
  ASSERT_TRUE(with_opt.is_array()) << with_opt.debug_to_string();
  ASSERT_EQ(with_opt.as_array_rows(), 2U);
  EXPECT_DOUBLE_EQ(plain[0], Cell(with_opt, 0, 1).as_number());
  EXPECT_DOUBLE_EQ(plain[1], Cell(with_opt, 1, 1).as_number());
}

TEST(GroupBy, OmittedParamIsBoundForIsomitted) {
  // The unsupplied `[u]` must read as omitted inside the body, not merely
  // be tolerated at the call boundary.
  const Value v =
      EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {10;20;30}, LAMBDA(v, [u], IF(ISOMITTED(u), SUM(v), -1)), 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 40.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 20.0);
}

TEST(GroupBy, RangeShapedBindingSurvivesWithOptional) {
  // `SUM(v)` must still see the group slice as a range, not an opaque
  // array, when the lambda also declares an optional param.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"A\";\"B\"}, {10;30;20}, LAMBDA(v, [u], AVERAGE(v)), 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 20.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 20.0);
}

TEST(GroupBy, LetBoundLambdaWithOptionalIsAccepted) {
  // Form B resolves the aggregator through the name environment, a
  // separate branch from the inline literal; it must agree on callability.
  const Value v = EvalSrc("=LET(f, LAMBDA(v, [u], SUM(v)), GROUPBY({\"A\";\"B\";\"A\"}, {10;20;30}, f, 0, 0))");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 40.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 20.0);
}

TEST(GroupBy, LetBoundLambdaBelowRequiredArityStillRejected) {
  // A lambda requiring three args is not callable through Form B either.
  const Value v = EvalSrc("=LET(f, LAMBDA(a, b, c, a+b+c), GROUPBY({\"A\";\"B\"}, {1;2}, f, 0, 0))");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(GroupBy, ZeroParamLambdaStillRejected) {
  // The other edge: GROUPBY supplies one argument, so a lambda declaring
  // no params at all cannot take it.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {1;2}, LAMBDA(42), 0, 0)");
  ASSERT_TRUE(v.is_error()) << v.debug_to_string();
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(GroupBy, AllOptionalSingleParamLambdaIsAccepted) {
  // `LAMBDA([v], ...)` has required = 0 and param_count = 1, so the single
  // supplied argument sits inside the window and binds normally.
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\";\"A\"}, {10;20;30}, LAMBDA([v], SUM(v)), 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 40.0);
  EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 20.0);
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
// Aggregator results Excel refuses per cell (Mac Excel 365 ja-JP)
// ---------------------------------------------------------------------------

// Every array result is #VALUE!, whatever its shape; a LAMBDA result is #CALC!.
TEST(GroupBy, AggregatorReturningArrayYieldsValueInThatCell) {
  const Value v = EvalSrc("=GROUPBY({\"A\";\"B\"}, {1;2}, LAMBDA(v, SEQUENCE(1, 3)), 0, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array()->cols, 2U);
  ASSERT_TRUE(Cell(v, 0, 1).is_error());
  EXPECT_EQ(Cell(v, 0, 1).as_error(), ErrorCode::Value);
  ASSERT_TRUE(Cell(v, 1, 1).is_error());
  EXPECT_EQ(Cell(v, 1, 1).as_error(), ErrorCode::Value);
}

TEST(GroupBy, AggregatorArrayResultOfAnyShapeIsValue) {
  for (const char* src : {
           "=GROUPBY({\"a\";\"a\"}, {1;2}, LAMBDA(v, v),, 0)",
           "=GROUPBY({\"a\"}, {1}, LAMBDA(v, SEQUENCE(1)),, 0)",
           "=GROUPBY({\"a\"}, {1}, LAMBDA(v, v),, 0)",
           "=GROUPBY({\"a\"}, {1}, LAMBDA(v, IF(v, 1)),, 0)",
           "=GROUPBY({\"a\"}, {1}, LAMBDA(v, w, IF(v, w)),, 0)",
           "=GROUPBY({\"a\"}, {1}, IFERROR,, 0)",
           "=GROUPBY({\"a\"}, {1}, IF,, 0)",
           "=GROUPBY({\"a\"}, {0}, IF,, 0)",
           "=GROUPBY({\"a\";\"a\"}, {1;2}, ABS,, 0)",
           "=GROUPBY({\"a\"}, {1}, ABS,, 0)",
           "=GROUPBY({\"a\"}, {1}, LAMBDA(v, ABS(v)),, 0)",
       }) {
    const Value v = EvalSrc(src);
    ASSERT_TRUE(v.is_array()) << src << " " << v.debug_to_string();
    ASSERT_TRUE(Cell(v, 0, 1).is_error()) << src;
    EXPECT_EQ(Cell(v, 0, 1).as_error(), ErrorCode::Value) << src;
  }
}

TEST(GroupBy, AggregatorReturningLambdaIsCalc) {
  const Value v = EvalSrc("=GROUPBY({\"a\"}, {1}, LAMBDA(v, LAMBDA(x, x)),, 0)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_TRUE(Cell(v, 0, 1).is_error());
  EXPECT_EQ(Cell(v, 0, 1).as_error(), ErrorCode::Calc);
}

TEST(GroupBy, ScalarAggregatorResultsAreKept) {
  EXPECT_EQ(Cell(EvalSrc("=GROUPBY({\"a\"}, {1}, LAMBDA(v, INDEX(v, 1)),, 0)"), 0, 1).as_number(), 1.0);
  EXPECT_EQ(Cell(EvalSrc("=GROUPBY({\"a\"}, {1}, LAMBDA(v, IF(SUM(v), 1)),, 0)"), 0, 1).as_number(), 1.0);
  EXPECT_EQ(Cell(EvalSrc("=GROUPBY({\"a\"}, {1}, LAMBDA(v, w, SUM(v) + 100 * SUM(w)),, 0)"), 0, 1).as_number(), 101.0);
}

TEST(GroupBy, ArrayResultInTotalsIsValue) {
  const Value v = EvalSrc("=GROUPBY({\"a\";\"b\"}, {1;2}, LAMBDA(v, v),, 1)");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array()->rows, 3U);
  for (std::uint32_t r = 0; r < 3U; ++r) {
    ASSERT_TRUE(Cell(v, r, 1).is_error()) << r;
    EXPECT_EQ(Cell(v, r, 1).as_error(), ErrorCode::Value) << r;
  }
}

// Results from Excel (ja-JP), cell by cell.

TEST(GroupByExcelParity, GbFh0Hdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(3, 1, Value::number(3.0));
  const Value v = EvalSrcIn("=GROUPBY(A1:A4,B1:B4,SUM,0)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 4, 2,
                            {"t:a", "n:2.000000", "t:b", "n:4.000000", "t:k", "n:0.000000", "t:合計", "n:6.000000"});
}

TEST(GroupByExcelParity, GbFh0Nohdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(3, 1, Value::number(3.0));
  const Value v = EvalSrcIn("=GROUPBY(A2:A4,B2:B4,SUM,0)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 3, 2, {"t:a", "n:2.000000", "t:b", "n:4.000000", "t:合計", "n:6.000000"});
}

TEST(GroupByExcelParity, GbFh1Hdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(3, 1, Value::number(3.0));
  const Value v = EvalSrcIn("=GROUPBY(A1:A4,B1:B4,SUM,1)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 3, 2, {"t:a", "n:2.000000", "t:b", "n:4.000000", "t:合計", "n:6.000000"});
}

TEST(GroupByExcelParity, GbFh1Nohdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(3, 1, Value::number(3.0));
  const Value v = EvalSrcIn("=GROUPBY(A2:A4,B2:B4,SUM,1)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 3, 2, {"t:a", "n:2.000000", "t:b", "n:3.000000", "t:合計", "n:5.000000"});
}

TEST(GroupByExcelParity, GbFh2Hdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(3, 1, Value::number(3.0));
  const Value v = EvalSrcIn("=GROUPBY(A1:A4,B1:B4,SUM,2)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 5, 2,
                            {"t:行フィールド 1", "t:値 1", "t:a", "n:2.000000", "t:b", "n:4.000000", "t:k",
                             "n:0.000000", "t:合計", "n:6.000000"});
}

TEST(GroupByExcelParity, GbFh2Nohdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(3, 1, Value::number(3.0));
  const Value v = EvalSrcIn("=GROUPBY(A2:A4,B2:B4,SUM,2)", wb, wb.sheet(0));
  excel_parity::ExpectCells(
      v, 4, 2, {"t:行フィールド 1", "t:値 1", "t:a", "n:2.000000", "t:b", "n:4.000000", "t:合計", "n:6.000000"});
}

TEST(GroupByExcelParity, GbFh3Hdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(3, 1, Value::number(3.0));
  const Value v = EvalSrcIn("=GROUPBY(A1:A4,B1:B4,SUM,3)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 4, 2, {"t:k", "t:v", "t:a", "n:2.000000", "t:b", "n:4.000000", "t:合計", "n:6.000000"});
}

TEST(GroupByExcelParity, GbFh3Nohdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(3, 1, Value::number(3.0));
  const Value v = EvalSrcIn("=GROUPBY(A2:A4,B2:B4,SUM,3)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 4, 2,
                            {"t:b", "n:1.000000", "t:a", "n:2.000000", "t:b", "n:3.000000", "t:合計", "n:5.000000"});
}

TEST(GroupByExcelParity, GbAutoNumhdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(2.0));
  const Value v = EvalSrcIn("=GROUPBY(A1:A3,B1:B3,SUM)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 4, 2,
                            {"t:a", "n:1.000000", "t:b", "n:2.000000", "t:k", "n:5.000000", "t:合計", "n:8.000000"});
}

TEST(GroupByExcelParity, GbAutoTextvals) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("y"));
  const Value v = EvalSrcIn("=GROUPBY(A1:A3,B1:B3,COUNTA)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 4, 2,
                            {"t:a", "n:1.000000", "t:b", "n:1.000000", "t:k", "n:1.000000", "t:合計", "n:3.000000"});
}

TEST(GroupByExcelParity, GbAutoRowfieldNum) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(2.0));
  const Value v = EvalSrcIn("=GROUPBY(A1:A3,B1:B3,SUM)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 3, 2, {"n:2.000000", "n:1.000000", "n:3.000000", "n:2.000000", "t:合計", "n:3.000000"});
}

TEST(GroupByExcelParity, GbAuto2valcols) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("v1"));
  wb.sheet(0).set_cell_value(0, 2, Value::text("v2"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(10.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 2, Value::number(20.0));
  const Value v = EvalSrcIn("=GROUPBY(A1:A3,B1:C3,SUM)", wb, wb.sheet(0));
  excel_parity::ExpectCells(
      v, 3, 3,
      {"t:a", "n:1.000000", "n:10.000000", "t:b", "n:2.000000", "n:20.000000", "t:合計", "n:3.000000", "n:30.000000"});
}

TEST(GroupByExcelParity, GbAuto2valcolsMixed) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("v1"));
  wb.sheet(0).set_cell_value(0, 2, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(1, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(10.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(2, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 2, Value::number(20.0));
  const Value v = EvalSrcIn("=GROUPBY(A1:A3,B1:C3,SUM)", wb, wb.sheet(0));
  excel_parity::ExpectCells(
      v, 3, 3,
      {"t:a", "n:1.000000", "n:10.000000", "t:b", "n:2.000000", "n:20.000000", "t:合計", "n:3.000000", "n:30.000000"});
}

TEST(GroupByExcelParity, GbAuto2keycols) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k1"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("k2"));
  wb.sheet(0).set_cell_value(0, 2, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("y"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("z"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("y"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=GROUPBY(A1:B4,C1:C4,SUM)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 4, 3,
                            {"t:a", "t:y", "n:3.000000", "t:a", "t:z", "n:2.000000", "t:b", "t:y", "n:1.000000",
                             "t:合計", "t:", "n:6.000000"});
}

TEST(GroupByExcelParity, GbSort2keys) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("y"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("z"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("y"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=GROUPBY(A2:B4,C2:C4,SUM,0,0)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 3, 3,
                            {"t:a", "t:y", "n:3.000000", "t:a", "t:z", "n:2.000000", "t:b", "t:y", "n:1.000000"});
}

TEST(GroupByExcelParity, GbSort2keysTotal) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("y"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("z"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("y"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=GROUPBY(A2:B4,C2:C4,SUM,0,2)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 6, 3,
                            {"t:a", "t:y", "n:3.000000", "t:a", "t:z", "n:2.000000", "t:a", "t:", "n:5.000000", "t:b",
                             "t:y", "n:1.000000", "t:b", "t:", "n:1.000000", "t:総計", "t:", "n:6.000000"});
}

TEST(GroupByExcelParity, GbSortMixedTypes) {
  const Value v = EvalSrc("=GROUPBY({\"b\";2;TRUE;\"a\";1},{1;2;3;4;5},SUM,0,0)");
  excel_parity::ExpectCells(v, 5, 2,
                            {"n:1.000000", "n:5.000000", "n:2.000000", "n:2.000000", "t:a", "n:4.000000", "t:b",
                             "n:1.000000", "b:1", "n:3.000000"});
}

TEST(GroupByExcelParity, GbSortCase) {
  const Value v = EvalSrc("=GROUPBY({\"b\";\"B\";\"a\";\"A\"},{1;2;3;4},SUM,0,0)");
  excel_parity::ExpectCells(v, 2, 2, {"t:a", "n:7.000000", "t:b", "n:3.000000"});
}

TEST(PivotByExcelParity, PbFh0Hdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("r"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("c"));
  wb.sheet(0).set_cell_value(0, 2, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("p"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=PIVOTBY(A1:A4,B1:B4,C1:C4,SUM,0)", wb, wb.sheet(0));
  excel_parity::ExpectCells(
      v, 5, 5, {"t:",         "t:c",        "t:p",        "t:q",       "t:合計",     "t:r",        "n:0.000000",
                "t:",         "t:",         "n:0.000000", "t:x",       "t:",         "n:1.000000", "n:3.000000",
                "n:4.000000", "t:y",        "t:",         "t:",        "n:2.000000", "n:2.000000", "t:合計",
                "n:0.000000", "n:1.000000", "n:5.000000", "n:6.000000"});
}

TEST(PivotByExcelParity, PbFh0Nohdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("r"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("c"));
  wb.sheet(0).set_cell_value(0, 2, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("p"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,0)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 4, 4,
                            {"t:", "t:p", "t:q", "t:合計", "t:x", "n:1.000000", "n:3.000000", "n:4.000000", "t:y",
                             "t:", "n:2.000000", "n:2.000000", "t:合計", "n:1.000000", "n:5.000000", "n:6.000000"});
}

TEST(PivotByExcelParity, PbFh1Hdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("r"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("c"));
  wb.sheet(0).set_cell_value(0, 2, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("p"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=PIVOTBY(A1:A4,B1:B4,C1:C4,SUM,1)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 4, 4,
                            {"t:", "t:p", "t:q", "t:合計", "t:x", "n:1.000000", "n:3.000000", "n:4.000000", "t:y",
                             "t:", "n:2.000000", "n:2.000000", "t:合計", "n:1.000000", "n:5.000000", "n:6.000000"});
}

TEST(PivotByExcelParity, PbFh1Nohdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("r"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("c"));
  wb.sheet(0).set_cell_value(0, 2, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("p"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,1)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 4, 3,
                            {"t:", "t:q", "t:合計", "t:x", "n:3.000000", "n:3.000000", "t:y", "n:2.000000",
                             "n:2.000000", "t:合計", "n:5.000000", "n:5.000000"});
}

TEST(PivotByExcelParity, PbFh2Hdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("r"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("c"));
  wb.sheet(0).set_cell_value(0, 2, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("p"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=PIVOTBY(A1:A4,B1:B4,C1:C4,SUM,2)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 7, 5, {"t:",
                                      "t:列フィールド 1",
                                      "t:",
                                      "t:",
                                      "t:",
                                      "t:",
                                      "t:c",
                                      "t:p",
                                      "t:q",
                                      "t:合計",
                                      "t:行フィールド 1",
                                      "t:値 1",
                                      "t:値 1",
                                      "t:値 1",
                                      "t:値 1",
                                      "t:r",
                                      "n:0.000000",
                                      "t:",
                                      "t:",
                                      "n:0.000000",
                                      "t:x",
                                      "t:",
                                      "n:1.000000",
                                      "n:3.000000",
                                      "n:4.000000",
                                      "t:y",
                                      "t:",
                                      "t:",
                                      "n:2.000000",
                                      "n:2.000000",
                                      "t:合計",
                                      "n:0.000000",
                                      "n:1.000000",
                                      "n:5.000000",
                                      "n:6.000000"});
}

TEST(PivotByExcelParity, PbFh2Nohdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("r"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("c"));
  wb.sheet(0).set_cell_value(0, 2, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("p"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,2)", wb, wb.sheet(0));
  excel_parity::ExpectCells(
      v, 6, 4, {"t:",         "t:列フィールド 1", "t:",        "t:",     "t:",         "t:p",        "t:q",
                "t:合計",     "t:行フィールド 1", "t:値 1",    "t:値 1", "t:値 1",     "t:x",        "n:1.000000",
                "n:3.000000", "n:4.000000",       "t:y",       "t:",     "n:2.000000", "n:2.000000", "t:合計",
                "n:1.000000", "n:5.000000",       "n:6.000000"});
}

TEST(PivotByExcelParity, PbFh3Hdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("r"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("c"));
  wb.sheet(0).set_cell_value(0, 2, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("p"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=PIVOTBY(A1:A4,B1:B4,C1:C4,SUM,3)", wb, wb.sheet(0));
  excel_parity::ExpectCells(
      v, 6, 4, {"t:",  "t:c", "t:",         "t:",         "t:",     "t:p",        "t:q",        "t:合計",
                "t:r", "t:v", "t:v",        "t:v",        "t:x",    "n:1.000000", "n:3.000000", "n:4.000000",
                "t:y", "t:",  "n:2.000000", "n:2.000000", "t:合計", "n:1.000000", "n:5.000000", "n:6.000000"});
}

TEST(PivotByExcelParity, PbFh3Nohdr) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("r"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("c"));
  wb.sheet(0).set_cell_value(0, 2, Value::text("v"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("p"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,3)", wb, wb.sheet(0));
  excel_parity::ExpectCells(
      v, 6, 3,
      {"t:", "t:p", "t:", "t:", "t:q", "t:合計", "t:x", "n:1.000000", "n:1.000000", "t:x", "n:3.000000", "n:3.000000",
       "t:y", "n:2.000000", "n:2.000000", "t:合計", "n:5.000000", "n:5.000000"});
}

TEST(PivotByExcelParity, PbLitFh0) {
  const Value v = EvalSrc("=PIVOTBY({\"x\";\"y\";\"x\"},{\"p\";\"q\";\"q\"},{1;2;3},SUM,0)");
  excel_parity::ExpectCells(v, 4, 4,
                            {"t:", "t:p", "t:q", "t:合計", "t:x", "n:1.000000", "n:3.000000", "n:4.000000", "t:y",
                             "t:", "n:2.000000", "n:2.000000", "t:合計", "n:1.000000", "n:5.000000", "n:6.000000"});
}

TEST(PivotByExcelParity, PbRngFh0) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("p"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,0)", wb, wb.sheet(0));
  excel_parity::ExpectCells(v, 4, 4,
                            {"t:", "t:p", "t:q", "t:合計", "t:x", "n:1.000000", "n:3.000000", "n:4.000000", "t:y",
                             "t:", "n:2.000000", "n:2.000000", "t:合計", "n:1.000000", "n:5.000000", "n:6.000000"});
}

TEST(PivotByExcelParity, PbRngFh0Isblank) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("p"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=ISBLANK(INDEX(PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,0),1,1))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_boolean()) << v.debug_to_string();
  EXPECT_EQ(v.as_boolean(), false);
}

TEST(PivotByExcelParity, PbLitIsblank) {
  const Value v = EvalSrc("=ISBLANK(INDEX(PIVOTBY({\"x\";\"y\";\"x\"},{\"p\";\"q\";\"q\"},{1;2;3},SUM,0),1,1))");
  ASSERT_TRUE(v.is_boolean()) << v.debug_to_string();
  EXPECT_EQ(v.as_boolean(), false);
}

TEST(PivotByExcelParity, PbRngType) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("p"));
  wb.sheet(0).set_cell_value(1, 2, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(2, 2, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("q"));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3.0));
  const Value v = EvalSrcIn("=TYPE(INDEX(PIVOTBY(A2:A4,B2:B4,C2:C4,SUM,0),1,1))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(GroupByExcelParity, Gb2LitTotal) {
  const Value v = EvalSrc("=GROUPBY({\"b\",\"y\";\"a\",\"z\";\"a\",\"y\"},{1;2;3},SUM,0)");
  excel_parity::ExpectCells(v, 4, 3,
                            {"t:a", "t:y", "n:3.000000", "t:a", "t:z", "n:2.000000", "t:b", "t:y", "n:1.000000",
                             "t:合計", "t:", "n:6.000000"});
}

TEST(GroupByExcelParity, Gb2LitTotalType) {
  const Value v = EvalSrc("=TYPE(INDEX(GROUPBY({\"b\",\"y\";\"a\",\"z\";\"a\",\"y\"},{1;2;3},SUM,0),4,2))");
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(GroupByExcelParity, Gb2LitTotalLen) {
  const Value v = EvalSrc("=ISBLANK(INDEX(GROUPBY({\"b\",\"y\";\"a\",\"z\";\"a\",\"y\"},{1;2;3},SUM,0),4,2))");
  ASSERT_TRUE(v.is_boolean()) << v.debug_to_string();
  EXPECT_EQ(v.as_boolean(), false);
}

// ---------------------------------------------------------------------------
// Text collation of group keys
// ---------------------------------------------------------------------------

// Mac Excel `=ROWS(GROUPBY({a;b},{1;2},SUM,0,0))` per locale, copied from the
// `groupby` column of tools/collation/data/fixtures/eq_{ja_JP,en_US,zh_CN}.tsv.
struct CollationPair {
  const char* a;
  const char* b;
  double ja_jp;
  double en_us;
  double zh_cn;
};

constexpr std::array<CollationPair, 26> kExcelGroupCounts = {{
    {"\xC3\x9F", "ss", 1, 1, 1},                                        // U+00DF | U+0073 U+0073
    {"\xC4\xB3", "ij", 1, 1, 1},                                        // U+0133 | U+0069 U+006A
    {"\xE2\x91\xA0", "1", 1, 1, 1},                                     // U+2460 | U+0031
    {"\xC2\xB2", "2", 1, 1, 1},                                         // U+00B2 | U+0032
    {"\xEF\xBC\xA1", "A", 1, 1, 1},                                     // U+FF21 | U+0041
    {"\xEF\xBD\x81", "A", 1, 1, 1},                                     // U+FF41 | U+0041
    {"\xEF\xBD\xB1", "\xE3\x82\xA2", 1, 1, 1},                          // U+FF71 | U+30A2
    {"\xE3\x81\x81", "\xE3\x81\x82", 1, 1, 2},                          // U+3041 | U+3042
    {"\xE3\x82\xA2", "\xE3\x81\x82", 1, 1, 2},                          // U+30A2 | U+3042
    {"\xE3\x83\xB4", "\xE3\x82\x94", 1, 1, 2},                          // U+30F4 | U+3094
    {"\xE3\x81\x8B\xE3\x83\xBC", "\xE3\x81\x8B\xE3\x81\x82", 1, 2, 2},  // U+304B U+30FC | U+304B U+3042
    {"\xE2\x85\xA0", "i", 1, 1, 2},                                     // U+2160 | U+0069
    {"\xC3\x84", "\xC3\xA4", 1, 1, 1},                                  // U+00C4 | U+00E4
    {"A", "a", 1, 1, 1},                                                // U+0041 | U+0061
    {"\xEF\xBD\xB6\xEF\xBE\x9E", "\xE3\x82\xAC", 1, 1, 1},              // U+FF76 U+FF9E | U+30AC
    {"\xEF\xBC\x91", "1", 1, 1, 1},                                     // U+FF11 | U+0031
    {"\xE3\x80\x80", " ", 1, 1, 1},                                     // U+3000 | U+0020
    {"\xE3\x82\x9D", "\xE3\x83\xBD", 1, 2, 1},                          // U+309D | U+30FD
    {"\xD1\x91", "\xD0\x81", 1, 1, 1},                                  // U+0451 | U+0401
    {"\xCF\x82", "\xCF\x83", 1, 1, 1},                                  // U+03C2 | U+03C3
    {"\xC3\xA4", "a", 2, 2, 2},                                         // U+00E4 | U+0061
    {"\xC3\xA6", "ae", 2, 2, 1},                                        // U+00E6 | U+0061 U+0065
    {"ss", "\xC3\x9F", 1, 1, 1},                                        // U+0073 U+0073 | U+00DF
    {"\xE0\xB9\x80\xE0\xB8\x81", "\xE0\xB8\x81\xE0\xB8\xB2", 2, 2, 2},  // U+0E40 U+0E01 | U+0E01 U+0E32
    {"\xE3\x81\x8B\xE3\x82\x9D", "\xE3\x81\x8B\xE3\x81\x8B", 1, 2, 2},  // U+304B U+309D | U+304B U+304B
    {"\xE3\x82\xAC", "\xEF\xBD\xB6\xEF\xBE\x9E", 1, 1, 1},              // U+30AC | U+FF76 U+FF9E
}};

ExcelProfile MacProfile(ExcelLocale locale) {
  return ExcelProfile{ExcelHost::kMac365, locale};
}

std::string Pair(const char* a, const char* b) {
  return std::string("{\"") + a + "\";\"" + b + "\"}";
}

double EvalNumber(const std::string& src, ExcelProfile profile) {
  const Value v = EvalSrcWithProfile(src, profile);
  EXPECT_TRUE(v.is_number()) << src << " -> " << v.debug_to_string();
  return v.is_number() ? v.as_number() : -1.0;
}

std::string EvalText(const std::string& src, ExcelProfile profile) {
  const Value v = EvalSrcWithProfile(src, profile);
  EXPECT_TRUE(v.is_text()) << src << " -> " << v.debug_to_string();
  return v.is_text() ? std::string(v.as_text()) : std::string();
}

TEST(GroupByCollation, MacGroupCountsMatchExcel) {
  for (const CollationPair& pair : kExcelGroupCounts) {
    const std::string src = "=ROWS(GROUPBY(" + Pair(pair.a, pair.b) + ",{1;2},SUM,0,0))";
    EXPECT_EQ(EvalNumber(src, MacProfile(ExcelLocale::kJaJP)), pair.ja_jp) << "ja_JP " << src;
    EXPECT_EQ(EvalNumber(src, MacProfile(ExcelLocale::kEnUS)), pair.en_us) << "en_US " << src;
    EXPECT_EQ(EvalNumber(src, MacProfile(ExcelLocale::kZhCN)), pair.zh_cn) << "zh_CN " << src;
  }
}

TEST(GroupByCollation, MergedKeysSumIntoOneGroup) {
  // U+FF71 and U+30A2 form one group in every Mac locale.
  const Value v = EvalSrcWithProfile("=GROUPBY({\"\xEF\xBD\xB1\";\"\xE3\x82\xA2\"},{10;20},SUM,0,0)",
                                     MacProfile(ExcelLocale::kJaJP));
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 30.0);
}

// Distinct keys sort as Mac Excel's SORT orders the same pair (the
// colfn_sort_* goldens in tools/collation/data/oracle_draft/targets).
TEST(GroupByCollation, MacKeyOrderFollowsDynamicCollation) {
  struct OrderCase {
    ExcelLocale locale;
    const char* a;
    const char* b;
    const char* first;
  };
  const std::array<OrderCase, 7> cases = {{
      {ExcelLocale::kJaJP, "\xC3\xA4", "a", "a"},
      {ExcelLocale::kEnUS, "\xC3\xA4", "a", "a"},
      {ExcelLocale::kZhCN, "\xC3\xA4", "a", "a"},
      {ExcelLocale::kEnUS, "\xE3\x81\x8B\xE3\x83\xBC", "\xE3\x81\x8B\xE3\x81\x82", "\xE3\x81\x8B\xE3\x83\xBC"},
      {ExcelLocale::kZhCN, "\xE3\x81\x8B\xE3\x83\xBC", "\xE3\x81\x8B\xE3\x81\x82", "\xE3\x81\x8B\xE3\x81\x82"},
      {ExcelLocale::kJaJP, "\xE0\xB9\x80\xE0\xB8\x81", "\xE0\xB8\x81\xE0\xB8\xB2", "\xE0\xB8\x81\xE0\xB8\xB2"},
      {ExcelLocale::kZhCN, "\xE0\xB9\x80\xE0\xB8\x81", "\xE0\xB8\x81\xE0\xB8\xB2", "\xE0\xB9\x80\xE0\xB8\x81"},
  }};
  for (const OrderCase& c : cases) {
    const std::string src = "=INDEX(GROUPBY(" + Pair(c.a, c.b) + ",{1;2},SUM,0,0),1,1)";
    EXPECT_EQ(EvalText(src, MacProfile(c.locale)), c.first) << src;
  }
}

// Windows groups across case, width and expansions but keeps hiragana and
// katakana apart (win-365-ja_JP colfn_groupby_*).
TEST(GroupByCollation, WindowsJaJpGroups) {
  const ExcelProfile win = win_365_ja_jp_profile();
  EXPECT_EQ(EvalNumber("=ROWS(GROUPBY({\"A\";\"a\"},{1;2},SUM,0,0))", win), 1.0);
  EXPECT_EQ(EvalNumber("=ROWS(GROUPBY({\"\xEF\xBD\xB6\xEF\xBE\x9E\";\"\xE3\x82\xAC\"},{1;2},SUM,0,0))", win), 1.0);
  EXPECT_EQ(EvalNumber("=ROWS(GROUPBY({\"\xE3\x81\x82\";\"\xE3\x82\xA2\"},{1;2},SUM,0,0))", win), 2.0);
  EXPECT_EQ(EvalNumber("=ROWS(GROUPBY({\"\xEF\xBC\xA1\";\"A\"},{1;2},SUM,0,0))", win), 1.0);  // Ａ
  EXPECT_EQ(EvalNumber("=ROWS(GROUPBY({\"\xC3\x9F\";\"ss\"},{1;2},SUM,0,0))", win), 1.0);     // ß
}

// PIVOTBY merges row and column items at the same strength as GROUPBY, so
// the GROUPBY counts carry over: one header row (or column) plus
// one per group.
TEST(PivotByCollation, MacRowGroupCountsMatchExcel) {
  for (const CollationPair& pair : kExcelGroupCounts) {
    const std::string src = "=ROWS(PIVOTBY(" + Pair(pair.a, pair.b) + ",{\"X\";\"X\"},{1;2},SUM,0,0,,0))";
    EXPECT_EQ(EvalNumber(src, MacProfile(ExcelLocale::kJaJP)), 1 + pair.ja_jp) << "ja_JP " << src;
    EXPECT_EQ(EvalNumber(src, MacProfile(ExcelLocale::kEnUS)), 1 + pair.en_us) << "en_US " << src;
    EXPECT_EQ(EvalNumber(src, MacProfile(ExcelLocale::kZhCN)), 1 + pair.zh_cn) << "zh_CN " << src;
  }
}

TEST(PivotByCollation, MacColumnGroupCountsMatchExcel) {
  for (const CollationPair& pair : kExcelGroupCounts) {
    const std::string src = "=COLUMNS(PIVOTBY({\"X\";\"X\"}," + Pair(pair.a, pair.b) + ",{1;2},SUM,0,0,,0))";
    EXPECT_EQ(EvalNumber(src, MacProfile(ExcelLocale::kJaJP)), 1 + pair.ja_jp) << "ja_JP " << src;
    EXPECT_EQ(EvalNumber(src, MacProfile(ExcelLocale::kEnUS)), 1 + pair.en_us) << "en_US " << src;
    EXPECT_EQ(EvalNumber(src, MacProfile(ExcelLocale::kZhCN)), 1 + pair.zh_cn) << "zh_CN " << src;
  }
}

TEST(PivotByCollation, MacRowOrderFollowsDynamicCollation) {
  // zh_CN sorts かあ before かー (colfn_sort_choon).
  const std::string src =
      "=INDEX(PIVOTBY({\"\xE3\x81\x8B\xE3\x83\xBC\";\"\xE3\x81\x8B\xE3\x81\x82\"},{\"X\";\"X\"},{1;2},SUM,0,0,,0),2,1)";
  EXPECT_EQ(EvalText(src, MacProfile(ExcelLocale::kZhCN)), "\xE3\x81\x8B\xE3\x81\x82");
  EXPECT_EQ(EvalText(src, MacProfile(ExcelLocale::kEnUS)), "\xE3\x81\x8B\xE3\x83\xBC");
}

TEST(PivotByCollation, WindowsJaJpGroups) {
  const ExcelProfile win = win_365_ja_jp_profile();
  EXPECT_EQ(EvalNumber("=ROWS(PIVOTBY({\"A\";\"a\"},{\"X\";\"X\"},{1;2},SUM,0,0,,0))", win), 2.0);
  EXPECT_EQ(EvalNumber("=ROWS(PIVOTBY({\"\xE3\x81\x82\";\"\xE3\x82\xA2\"},{\"X\";\"X\"},{1;2},SUM,0,0,,0))", win), 3.0);
}

// A volatile aggregator argument must be evaluated once: the arity peek and
// the callable resolution would otherwise pick different LAMBDAs and disagree.
TEST(GroupBy, VolatileAggregatorArgumentEvaluatedOnce) {
  for (int i = 0; i < 200; ++i) {
    const Value v = EvalSrc("=GROUPBY({1;1;2},{1;2;3},IF(RAND()<0.5,LAMBDA(x,SUM(x)),LAMBDA(x,y,SUM(x))),,0)");
    ASSERT_TRUE(v.is_array()) << "iteration " << i << ": " << v.debug_to_string();
    ASSERT_EQ(v.as_array_rows(), 2U);
    ASSERT_EQ(v.as_array_cols(), 2U);
    EXPECT_DOUBLE_EQ(Cell(v, 0, 1).as_number(), 3.0) << "iteration " << i;
    EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 3.0) << "iteration " << i;
  }
}

TEST(PivotBy, VolatileAggregatorArgumentEvaluatedOnce) {
  for (int i = 0; i < 200; ++i) {
    const Value v = EvalSrc(
        "=PIVOTBY({1;1;2},{\"X\";\"X\";\"X\"},{1;2;3},IF(RAND()<0.5,LAMBDA(x,SUM(x)),LAMBDA(x,y,SUM(x))),0,0,,0)");
    ASSERT_TRUE(v.is_array()) << "iteration " << i << ": " << v.debug_to_string();
    // Column-header row, then one row per group.
    ASSERT_EQ(v.as_array_rows(), 3U);
    ASSERT_EQ(v.as_array_cols(), 2U);
    EXPECT_DOUBLE_EQ(Cell(v, 1, 1).as_number(), 3.0) << "iteration " << i;
    EXPECT_DOUBLE_EQ(Cell(v, 2, 1).as_number(), 3.0) << "iteration " << i;
  }
}

using excel_parity::ExactCell;

// field_relationship 0 (or omitted) keeps the key columns a hierarchy: a value
// sort ranks each outer group by its aggregate; an omitted filter keeps every row.
TEST(GroupByFieldRelationship, HierarchyLayoutsMatchExcel) {
  const std::vector<excel_parity::ExactCase> cases = {
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,,,0)",
       5,
       3,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"),
        ExactCell(1.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,1,,0)",
       5,
       3,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"),
        ExactCell(1.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,-1,,0)",
       5,
       3,
       {ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"),
        ExactCell("q"), ExactCell(1.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"),
        ExactCell(4.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,2,,0)",
       5,
       3,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"),
        ExactCell(1.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,-2,,0)",
       5,
       3,
       {ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("b"),
        ExactCell("q"), ExactCell(1.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("a"),
        ExactCell(3.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,3,,0)",
       5,
       3,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"),
        ExactCell("q"), ExactCell(1.0), ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"),
        ExactCell(5.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,-3,,0)",
       5,
       3,
       {ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"),
        ExactCell("q"), ExactCell(1.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("a"), ExactCell("q"),
        ExactCell(2.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,{1,2},,0)",
       5,
       3,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"),
        ExactCell(1.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,{2,1},,0)",
       5,
       3,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"),
        ExactCell(1.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,{-2,1},,0)",
       5,
       3,
       {ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("b"),
        ExactCell("q"), ExactCell(1.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("a"),
        ExactCell(3.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,{2,-1},,0)",
       5,
       3,
       {ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"),
        ExactCell("q"), ExactCell(1.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"),
        ExactCell(4.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,1,,,0)",
       6,
       3,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"),
        ExactCell(1.0), ExactCell("合計"), ExactCell(""), ExactCell(15.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,-1,,,0)",
       6,
       3,
       {ExactCell("合計"), ExactCell(""), ExactCell(15.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0),
        ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"),
        ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"), ExactCell(1.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,2,,,0)",
       8,
       3,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"),    ExactCell("z"), ExactCell(4.0),
        ExactCell("a"), ExactCell(""),  ExactCell(6.0), ExactCell("b"),    ExactCell("a"), ExactCell(3.0),
        ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"),    ExactCell("q"), ExactCell(1.0),
        ExactCell("b"), ExactCell(""),  ExactCell(9.0), ExactCell("総計"), ExactCell(""),  ExactCell(15.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,-2,,,0)",
       8,
       3,
       {ExactCell("総計"), ExactCell(""),  ExactCell(15.0), ExactCell("a"), ExactCell(""),  ExactCell(6.0),
        ExactCell("a"),    ExactCell("q"), ExactCell(2.0),  ExactCell("a"), ExactCell("z"), ExactCell(4.0),
        ExactCell("b"),    ExactCell(""),  ExactCell(9.0),  ExactCell("b"), ExactCell("a"), ExactCell(3.0),
        ExactCell("b"),    ExactCell("p"), ExactCell(5.0),  ExactCell("b"), ExactCell("q"), ExactCell(1.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,3,1,,,0)",
       6,
       3,
       {ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"),
        ExactCell("z"), ExactCell(4.0), ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("q"),
        ExactCell(1.0), ExactCell("合計"), ExactCell(""), ExactCell(10.0)}},
      {"GROUPBY({\"H1\",\"H2\";\"b\",\"p\";\"a\",\"q\";\"b\",\"p\"},{\"V\";1;2;3},SUM,3,1,2,,0)",
       4,
       3,
       {ExactCell("H1"), ExactCell("H2"), ExactCell("V"), ExactCell("a"), ExactCell("q"), ExactCell(2.0),
        ExactCell("b"), ExactCell("p"), ExactCell(4.0), ExactCell("合計"), ExactCell(""), ExactCell(6.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5,1;2,2;3,3;4,4;1,5},SUM,0,0,4,,0)",
       5,
       4,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0),
        ExactCell(4.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell(1.0), ExactCell("b"), ExactCell("a"),
        ExactCell(3.0), ExactCell(3.0), ExactCell("b"), ExactCell("q"), ExactCell(1.0), ExactCell(5.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,,{TRUE;FALSE;TRUE;"
       "TRUE;TRUE},0)",
       4,
       3,
       {ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"),
        ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"), ExactCell(1.0)}},
      {"GROUPBY({\"b\";\"a\";\"b\"},{1;2;3},SUM,0,1,-2,,0)",
       3,
       2,
       {ExactCell("b"), ExactCell(4.0), ExactCell("a"), ExactCell(2.0), ExactCell("合計"), ExactCell(6.0)}},
      {"GROUPBY({\"A\";\"B\";\"A\"},{10;20;30},SUM,0,0,,,0)",
       2,
       2,
       {ExactCell("A"), ExactCell(40.0), ExactCell("B"), ExactCell(20.0)}},
      {"GROUPBY({\"A\";\"B\";\"A\"},{10;20;30},SUM,0,0,,,)",
       2,
       2,
       {ExactCell("A"), ExactCell(40.0), ExactCell("B"), ExactCell(20.0)}},
      {"GROUPBY({\"x\",\"p\";\"y\",\"q\";\"x\",\"q\";\"y\",\"p\"},{1;2;3;4},SUM,0,2,,,0)",
       7,
       3,
       {ExactCell("x"),    ExactCell("p"), ExactCell(1.0), ExactCell("x"), ExactCell("q"), ExactCell(3.0),
        ExactCell("x"),    ExactCell(""),  ExactCell(4.0), ExactCell("y"), ExactCell("p"), ExactCell(4.0),
        ExactCell("y"),    ExactCell("q"), ExactCell(2.0), ExactCell("y"), ExactCell(""),  ExactCell(6.0),
        ExactCell("総計"), ExactCell(""),  ExactCell(10.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\"},{1;2;3;4},SUM,0,0,2,,0)",
       4,
       3,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(1.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\"},{1;2;3;4},SUM,0,0,3,,0)",
       4,
       3,
       {ExactCell("b"), ExactCell("p"), ExactCell(1.0), ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("a"),
        ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0)}},
      {"GROUPBY({\"A\";\"B\";\"A\"},{10;20;30},SUM,0,0,,)",
       2,
       2,
       {ExactCell("A"), ExactCell(40.0), ExactCell("B"), ExactCell(20.0)}},
      {"GROUPBY({\"x\",\"p\";\"x\",\"q\";\"y\",\"p\";\"y\",\"q\";\"x\",\"p\"},{1;2;3;4;5},PERCENTOF,0,2)",
       7,
       3,
       {ExactCell("x"),    ExactCell("p"), ExactCell(0.4),
        ExactCell("x"),    ExactCell("q"), ExactCell(0.13333333333333333),
        ExactCell("x"),    ExactCell(""),  ExactCell(0.5333333333333333),
        ExactCell("y"),    ExactCell("p"), ExactCell(0.2),
        ExactCell("y"),    ExactCell("q"), ExactCell(0.26666666666666666),
        ExactCell("y"),    ExactCell(""),  ExactCell(0.4666666666666667),
        ExactCell("総計"), ExactCell(""),  ExactCell(1.0)}},
      {"GROUPBY({\"x\",\"p\";\"x\",\"q\";\"y\",\"p\";\"y\",\"q\";\"x\",\"p\"},{1;2;3;4;5},LAMBDA(a,b,ROWS(b)),0,2)",
       7,
       3,
       {ExactCell("x"),    ExactCell("p"), ExactCell(5.0), ExactCell("x"), ExactCell("q"), ExactCell(5.0),
        ExactCell("x"),    ExactCell(""),  ExactCell(5.0), ExactCell("y"), ExactCell("p"), ExactCell(5.0),
        ExactCell("y"),    ExactCell("q"), ExactCell(5.0), ExactCell("y"), ExactCell(""),  ExactCell(5.0),
        ExactCell("総計"), ExactCell(""),  ExactCell(5.0)}},
  };
  for (const excel_parity::ExactCase& c : cases) {
    SCOPED_TRACE(c.formula);
    excel_parity::ExpectExactCells(EvalSrc(c.formula), c);
  }
}

// field_relationship 1 sorts the leaf rows as a flat table; ties fall to the
// other key columns ascending.
TEST(GroupByFieldRelationship, FlatTableLayoutsMatchExcel) {
  const std::vector<excel_parity::ExactCase> cases = {
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,,,1)",
       5,
       3,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"),
        ExactCell(1.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,1,,1)",
       5,
       3,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"),
        ExactCell(1.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,-1,,1)",
       5,
       3,
       {ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"),
        ExactCell("q"), ExactCell(1.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"),
        ExactCell(4.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,2,,1)",
       5,
       3,
       {ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("a"),
        ExactCell("q"), ExactCell(2.0), ExactCell("b"), ExactCell("q"), ExactCell(1.0), ExactCell("a"), ExactCell("z"),
        ExactCell(4.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,-2,,1)",
       5,
       3,
       {ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("b"),
        ExactCell("q"), ExactCell(1.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("a"),
        ExactCell(3.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,3,,1)",
       5,
       3,
       {ExactCell("b"), ExactCell("q"), ExactCell(1.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"), ExactCell("p"),
        ExactCell(5.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,-3,,1)",
       5,
       3,
       {ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("b"), ExactCell("q"),
        ExactCell(1.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,{1,2},,1)",
       5,
       3,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"),
        ExactCell(1.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,{2,1},,1)",
       5,
       3,
       {ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("a"),
        ExactCell("q"), ExactCell(2.0), ExactCell("b"), ExactCell("q"), ExactCell(1.0), ExactCell("a"), ExactCell("z"),
        ExactCell(4.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,{-2,1},,1)",
       5,
       3,
       {ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("b"),
        ExactCell("q"), ExactCell(1.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("a"),
        ExactCell(3.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,{2,-1},,1)",
       5,
       3,
       {ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"),
        ExactCell("q"), ExactCell(1.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"),
        ExactCell(4.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,1,,,1)",
       6,
       3,
       {ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"),
        ExactCell(1.0), ExactCell("合計"), ExactCell(""), ExactCell(15.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,-1,,,1)",
       6,
       3,
       {ExactCell("合計"), ExactCell(""), ExactCell(15.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0),
        ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"),
        ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"), ExactCell(1.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,3,1,,,1)",
       6,
       3,
       {ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("a"),
        ExactCell("z"), ExactCell(4.0), ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("q"),
        ExactCell(1.0), ExactCell("合計"), ExactCell(""), ExactCell(10.0)}},
      {"GROUPBY({\"H1\",\"H2\";\"b\",\"p\";\"a\",\"q\";\"b\",\"p\"},{\"V\";1;2;3},SUM,3,1,2,,1)",
       4,
       3,
       {ExactCell("H1"), ExactCell("H2"), ExactCell("V"), ExactCell("b"), ExactCell("p"), ExactCell(4.0),
        ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("合計"), ExactCell(""), ExactCell(6.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5,1;2,2;3,3;4,4;1,5},SUM,0,0,4,,1)",
       5,
       4,
       {ExactCell("b"), ExactCell("p"), ExactCell(5.0), ExactCell(1.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0),
        ExactCell(2.0), ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell(3.0), ExactCell("a"), ExactCell("z"),
        ExactCell(4.0), ExactCell(4.0), ExactCell("b"), ExactCell("q"), ExactCell(1.0), ExactCell(5.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,,{TRUE;FALSE;TRUE;"
       "TRUE;TRUE},1)",
       4,
       3,
       {ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"),
        ExactCell("p"), ExactCell(5.0), ExactCell("b"), ExactCell("q"), ExactCell(1.0)}},
      {"GROUPBY({\"b\";\"a\";\"b\"},{1;2;3},SUM,0,1,-2,,1)",
       3,
       2,
       {ExactCell("b"), ExactCell(4.0), ExactCell("a"), ExactCell(2.0), ExactCell("合計"), ExactCell(6.0)}},
      {"GROUPBY({\"A\";\"B\";\"A\"},{10;20;30},SUM,0,0,,,1)",
       2,
       2,
       {ExactCell("A"), ExactCell(40.0), ExactCell("B"), ExactCell(20.0)}},
      {"GROUPBY({\"x\",\"p\";\"y\",\"q\";\"x\",\"q\";\"y\",\"p\"},{1;2;3;4},SUM,0,0,,,1)",
       4,
       3,
       {ExactCell("x"), ExactCell("p"), ExactCell(1.0), ExactCell("x"), ExactCell("q"), ExactCell(3.0), ExactCell("y"),
        ExactCell("p"), ExactCell(4.0), ExactCell("y"), ExactCell("q"), ExactCell(2.0)}},
      {"GROUPBY({\"x\",\"p\";\"y\",\"q\";\"x\",\"q\";\"y\",\"p\"},{1;2;3;4},SUM,0,1,,,1)",
       5,
       3,
       {ExactCell("x"), ExactCell("p"), ExactCell(1.0), ExactCell("x"), ExactCell("q"), ExactCell(3.0), ExactCell("y"),
        ExactCell("p"), ExactCell(4.0), ExactCell("y"), ExactCell("q"), ExactCell(2.0), ExactCell("合計"),
        ExactCell(""), ExactCell(10.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\"},{1;2;3;4},SUM,0,0,2,,1)",
       4,
       3,
       {ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(1.0), ExactCell("a"),
        ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\"},{1;2;3;4},SUM,0,0,-2,,1)",
       4,
       3,
       {ExactCell("a"), ExactCell("z"), ExactCell(4.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("b"),
        ExactCell("p"), ExactCell(1.0), ExactCell("b"), ExactCell("a"), ExactCell(3.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\"},{1;2;3;4},SUM,0,0,3,,1)",
       4,
       3,
       {ExactCell("b"), ExactCell("p"), ExactCell(1.0), ExactCell("a"), ExactCell("q"), ExactCell(2.0), ExactCell("b"),
        ExactCell("a"), ExactCell(3.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0)}},
      {"GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\"},{1;2;3;4},SUM,0,0,-1,,1)",
       4,
       3,
       {ExactCell("b"), ExactCell("a"), ExactCell(3.0), ExactCell("b"), ExactCell("p"), ExactCell(1.0), ExactCell("a"),
        ExactCell("q"), ExactCell(2.0), ExactCell("a"), ExactCell("z"), ExactCell(4.0)}},
  };
  for (const excel_parity::ExactCase& c : cases) {
    SCOPED_TRACE(c.formula);
    excel_parity::ExpectExactCells(EvalSrc(c.formula), c);
  }
}

TEST(GroupByFieldRelationship, RejectsOutOfDomainValuesAndFlatSubtotals) {
  const char* const formulas[] = {
      "GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,{3,1},,0)",
      "GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,0,{3,1},,1)",
      "GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,2,,,1)",
      "GROUPBY({\"b\",\"p\";\"a\",\"q\";\"b\",\"a\";\"a\",\"z\";\"b\",\"q\"},{5;2;3;4;1},SUM,0,-2,,,1)",
      "GROUPBY({\"A\";\"B\";\"A\"},{10;20;30},SUM,0,0,,,2)",
      "GROUPBY({\"A\";\"B\";\"A\"},{10;20;30},SUM,0,0,,,-1)",
      "GROUPBY({\"A\";\"B\";\"A\"},{10;20;30},SUM,0,0,,,\"x\")",
      "GROUPBY({\"x\",\"p\";\"y\",\"q\";\"x\",\"q\";\"y\",\"p\"},{1;2;3;4},SUM,0,2,,,1)",
      "GROUPBY({\"F1\",\"F2\";\"x\",\"p\";\"y\",\"q\";\"x\",\"q\"},{\"V\";1;2;3},SUM,3,2,,,1)",
  };
  for (const char* formula : formulas) {
    const Value v = EvalSrc(formula);
    ASSERT_TRUE(v.is_error()) << formula << ": " << v.debug_to_string();
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << formula;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
