//
// End-to-end tests for the rank / percentile-rank family:
// RANK (legacy), RANK.EQ, RANK.AVG, PERCENTRANK (legacy),
// PERCENTRANK.INC, PERCENTRANK.EXC.
//
// All six ride the lazy-dispatch seam because the array argument must
// retain its range / Ref / ArrayLiteral shape (the eager path would
// flatten it alongside the trailing scalar).
//
// Like the other descriptive-stats functions (MEDIAN / LARGE / SMALL),
// the array argument is filtered to numeric cells only — Text, Bool,
// and Blank are skipped silently.

#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;
using formulon::test::EvalSourceIn;

// ---------------------------------------------------------------------------
// RANK / RANK.EQ
// ---------------------------------------------------------------------------

TEST(BuiltinsRank, DescendingBasic) {
  // Array {30, 20, 10}: 20 has 1 strictly greater -> rank 2.
  const Value v = EvalSource("=RANK(20, {30;20;10}, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsRank, PercentRankExcessiveSignificanceIsNum) {
  for (const char* formula : {"=PERCENTRANK.INC({1;2}, 1, 309)", "=PERCENTRANK.EXC({1;2}, 1, 309)"}) {
    const Value v = EvalSource(formula);
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Num) << formula;
  }
}

TEST(BuiltinsRank, AscendingBasic) {
  // Array {30, 20, 10} ascending: 20 has 1 strictly less (10) -> rank 2.
  const Value v = EvalSource("=RANK(20, {30;20;10}, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsRank, TiesShareRank) {
  // Array {30, 20, 20, 10} descending: both 20s share rank 2
  // (one value strictly greater: 30).
  const Value v = EvalSource("=RANK(20, {30;20;20;10}, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsRank, OrderOmittedIsDescending) {
  const Value v = EvalSource("=RANK(30, {30;20;10})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsRank, NotPresentIsNA) {
  const Value v = EvalSource("=RANK(15, {30;20;10})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsRank, AllNonNumericRangeIsNA) {
  // A range whose cells are all text filters down to empty -> #N/A.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("b"));
  const Value v = EvalSourceIn("=RANK(1, A1:A2, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsRank, NonNumericNumberIsValue) {
  const Value v = EvalSource("=RANK(\"x\", {1;2;3})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsRank, ErrorInRangePropagates) {
  // A cell carrying #DIV/0! inside the range is propagated.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::error(ErrorCode::Div0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  const Value v = EvalSourceIn("=RANK(10, A1:A3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsRankEq, AliasMatchesRank) {
  const Value v = EvalSource("=RANK.EQ(20, {30;20;20;10}, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// RANK.AVG
// ---------------------------------------------------------------------------

TEST(BuiltinsRankAvg, NoTiesMatchesEq) {
  const Value v = EvalSource("=RANK.AVG(20, {30;20;10}, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsRankAvg, TiesAveraged) {
  // {30, 20, 20, 10} descending: two 20s occupy rank slots 2 and 3,
  // average 2.5.
  const Value v = EvalSource("=RANK.AVG(20, {30;20;20;10}, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.5);
}

TEST(BuiltinsRankAvg, FivewayTieAveraged) {
  // Five equal values occupy rank slots 1..5; average = 3.
  const Value v = EvalSource("=RANK.AVG(5, {5;5;5;5;5}, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsRankAvg, AscendingTies) {
  // {1, 2, 2, 3} ascending: two 2s occupy slots 2 and 3, average 2.5.
  const Value v = EvalSource("=RANK.AVG(2, {1;2;2;3}, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.5);
}

// ---------------------------------------------------------------------------
// PERCENTRANK / PERCENTRANK.INC
// ---------------------------------------------------------------------------

TEST(BuiltinsPercentRankInc, MidpointBasic) {
  // Array {1..5}, x = 3 (index 2 of 5). raw = 2/4 = 0.5.
  const Value v = EvalSource("=PERCENTRANK.INC({1;2;3;4;5}, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5);
}

TEST(BuiltinsPercentRankInc, MinIsZero) {
  const Value v = EvalSource("=PERCENTRANK.INC({10;20;30}, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsPercentRankInc, MaxIsOne) {
  const Value v = EvalSource("=PERCENTRANK.INC({10;20;30}, 30)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsPercentRankInc, InterpolatedDefaultSignificance) {
  // Sorted {1, 2, 4}: x = 3 lies between indices 1 and 2.
  // raw = (1 + (3 - 2)/(4 - 2)) / (3 - 1) = 1.5 / 2 = 0.75. Default
  // significance is 3 so 0.75 is exact.
  const Value v = EvalSource("=PERCENTRANK.INC({1;2;4}, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.75);
}

TEST(BuiltinsPercentRankInc, SignificanceTwoTruncatesNotRounds) {
  // Sorted {1..5}, x = 2.5. raw = (1 + 0.5) / 4 = 0.375.
  // Truncated (not rounded) to 2 digits -> 0.37.
  const Value v = EvalSource("=PERCENTRANK.INC({1;2;3;4;5}, 2.5, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.37);
}

TEST(BuiltinsPercentRankInc, OutOfRangeIsNA) {
  const Value v = EvalSource("=PERCENTRANK.INC({1;2;3}, 4)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsPercentRankInc, NonNumericXIsValue) {
  const Value v = EvalSource("=PERCENTRANK.INC({1;2;3}, \"x\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsPercentRank, LegacyMatchesInc) {
  const Value v = EvalSource("=PERCENTRANK({1;2;3;4;5}, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5);
}

TEST(BuiltinsPercentRankInc, SingleElementIsNA) {
  const Value v = EvalSource("=PERCENTRANK.INC({5}, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsPercentRankInc, SignificanceZeroIsNum) {
  const Value v = EvalSource("=PERCENTRANK.INC({1;2;3;4;5}, 3, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsPercentRankInc, SignificanceNegativeIsNum) {
  const Value v = EvalSource("=PERCENTRANK.INC({1;2;3;4;5}, 3, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// PERCENTRANK.EXC
// ---------------------------------------------------------------------------

TEST(BuiltinsPercentRankExc, MidpointBasic) {
  // Sorted {1..5}, x = 3 at index 2. raw = (2 + 1) / (5 + 1) = 0.5.
  const Value v = EvalSource("=PERCENTRANK.EXC({1;2;3;4;5}, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5);
}

TEST(BuiltinsPercentRankExc, AtMinIsOneOverNPlus1) {
  // x = 1 at index 0. raw = 1/6 ~ 0.1666..., truncated to 3 digits 0.166.
  const Value v = EvalSource("=PERCENTRANK.EXC({1;2;3;4;5}, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.166);
}

TEST(BuiltinsPercentRankExc, AtMaxIsNOverNPlus1) {
  // x = 5 at index 4. raw = 5/6 ~ 0.8333..., truncated to 3 digits 0.833.
  const Value v = EvalSource("=PERCENTRANK.EXC({1;2;3;4;5}, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.833);
}

TEST(BuiltinsPercentRankExc, OutOfRangeIsNA) {
  const Value v = EvalSource("=PERCENTRANK.EXC({1;2;3;4;5}, 6)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Registry pins
// ---------------------------------------------------------------------------
//
// All six names are lazy-dispatched, so `registry.lookup()` returns
// nullptr for them. We instead assert end-to-end evaluation works,
// which proves the dispatch table wiring.

TEST(BuiltinsRankRegistry, AllNamesDispatch) {
  for (const char* formula : {
           "=RANK(1, {1;2;3}, 0)",
           "=RANK.EQ(1, {1;2;3}, 0)",
           "=RANK.AVG(1, {1;2;3}, 0)",
           "=PERCENTRANK({1;2;3}, 2)",
           "=PERCENTRANK.INC({1;2;3}, 2)",
           "=PERCENTRANK.EXC({1;2;3}, 2)",
       }) {
    const Value v = EvalSource(formula);
    EXPECT_FALSE(v.is_error() && v.as_error() == ErrorCode::Name) << "unrouted: " << formula;
  }
}

TEST(BuiltinsRank, PercentRankInterpolatesPastADuplicateRun) {
  // Sorted {-4,-1,3,7.5,12,12,56}: 40 lies between the second 12 (index 5)
  // and 56, so the anchor is the run's last index.
  const Value inc = EvalSource("=PERCENTRANK.INC({12;-1;7.5;56;12;-4;3}, 40)");
  ASSERT_TRUE(inc.is_number()) << inc.debug_to_string();
  EXPECT_DOUBLE_EQ(inc.as_number(), 0.939);
  const Value exc = EvalSource("=PERCENTRANK.EXC({12;-1;7.5;56;12;-4;3}, 40)");
  ASSERT_TRUE(exc.is_number()) << exc.debug_to_string();
  EXPECT_DOUBLE_EQ(exc.as_number(), 0.829);
}

// PERCENTRANK's x and significance coerce as direct arguments, and an array
// in either evaluates per element (Mac Excel 365).
TEST(BuiltinsRank, PercentRankCoercesXAndSignificance) {
  struct Case {
    const char* source;
    double expected;
  };
  for (const Case& c :
       {Case{"=PERCENTRANK({0,1,2},TRUE)", 0.5}, Case{"=PERCENTRANK.INC({10,20,30,40},\"20\")", 0.333},
        Case{"=PERCENTRANK.INC({10,20,30,40},25,TRUE)", 0.5}, Case{"=PERCENTRANK.EXC({0,1,2},TRUE)", 0.5}}) {
    const Value v = EvalSource(c.source);
    ASSERT_TRUE(v.is_number()) << c.source << ": " << v.debug_to_string();
    EXPECT_DOUBLE_EQ(v.as_number(), c.expected) << c.source;
  }
  const Value text = EvalSource("=PERCENTRANK({0,1,2},\"x\")");
  ASSERT_TRUE(text.is_error());
  EXPECT_EQ(text.as_error(), ErrorCode::Value);
  const Value zero_sig = EvalSource("=PERCENTRANK({0,1,2},1,0)");
  ASSERT_TRUE(zero_sig.is_error());
  EXPECT_EQ(zero_sig.as_error(), ErrorCode::Num);
}

TEST(BuiltinsRank, PercentRankArrayXEvaluatesPerElement) {
  const Value v = EvalSource("=PERCENTRANK.INC({10,20,30,40,50,60,70},{10,35})");
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array()->cols, 2U);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[0].as_number(), 0.0);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[1].as_number(), 0.416);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
