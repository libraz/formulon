// Conditional-format evaluator tests: population predicates.

#include <cstdint>
#include <initializer_list>
#include <vector>

#include "cf/cf_evaluator.h"
#include "cf/cf_helpers.h"
#include "cf/cf_types.h"
#include "cf_evaluator_fixtures.h"
#include "gtest/gtest.h"
#include "utils/error.h"
#include "value.h"
namespace formulon::cf {
namespace {

using test::At;
using test::CFEvalHarness;
using test::MakeRange;
using test::MakeRule;
using test::PopulateLinearPopulation;
using test::RGB;
using test::SqrefContext;
using test::TwoStopMinMax;

// ---------------------------------------------------------------------------
// DuplicateValues / UniqueValues
// ---------------------------------------------------------------------------

TEST(CFEvaluator, DuplicateValuesWithoutSqrefDoesNotMatch) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(7.0));
  harness.sheet.set_cell_value(0, 1, Value::number(7.0));
  CFRule r = MakeRule(RuleType::DuplicateValues);
  // ctx.sqref intentionally nullptr.
  EXPECT_FALSE(match_rule(r, Value::number(7.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, DuplicateValuesMatchesNumberAppearingTwice) {
  CFEvalHarness harness;
  // A1:A3 = [7, 7, 9]. Value 7 appears twice → duplicate; 9 appears once.
  harness.sheet.set_cell_value(0, 0, Value::number(7.0));
  harness.sheet.set_cell_value(1, 0, Value::number(7.0));
  harness.sheet.set_cell_value(2, 0, Value::number(9.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(7.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(9.0), ctx));
}

TEST(CFEvaluator, DuplicateValuesNumbersUseExactEquality) {
  // 1.0 and 1.0000000001 are distinct under IEEE-754 equality, even
  // though they round to the same display string.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(1.0));
  harness.sheet.set_cell_value(1, 0, Value::number(1.0000000001));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 1, 0)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::number(1.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(1.0000000001), ctx));
}

TEST(CFEvaluator, DuplicateValuesTextIsAsciiCaseInsensitive) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::text("apple"));
  harness.sheet.set_cell_value(1, 0, Value::text("APPLE"));
  harness.sheet.set_cell_value(2, 0, Value::text("orange"));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::text("apple"), ctx));
  EXPECT_TRUE(match_rule(r, Value::text("Apple"), ctx));
  EXPECT_FALSE(match_rule(r, Value::text("orange"), ctx));
}

TEST(CFEvaluator, DuplicateValuesBooleansMatchByIdentity) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::boolean(true));
  harness.sheet.set_cell_value(1, 0, Value::boolean(true));
  harness.sheet.set_cell_value(2, 0, Value::boolean(false));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::boolean(true), ctx));
  EXPECT_FALSE(match_rule(r, Value::boolean(false), ctx));
}

TEST(CFEvaluator, DuplicateValuesIsCrossKindFalse) {
  // Number 1 and text "1" are distinct, even though Excel coerces in
  // some contexts. Mirroring the cellIs cross-kind=false stance.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(1.0));
  harness.sheet.set_cell_value(1, 0, Value::text("1"));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 1, 0)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::number(1.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::text("1"), ctx));
}

TEST(CFEvaluator, DuplicateValuesErrorsAndBlanksDoNotMatch) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::error(ErrorCode::NA));
  harness.sheet.set_cell_value(1, 0, Value::error(ErrorCode::NA));
  // A3 left blank.
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
}

TEST(CFEvaluator, DuplicateValuesAcrossMultipleSqrefRanges) {
  // sqref unions A1:A2 and C1:C2. Value 5 appears in A1 and C2 → dup.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(5.0));
  harness.sheet.set_cell_value(1, 0, Value::number(99.0));
  harness.sheet.set_cell_value(0, 2, Value::number(11.0));
  harness.sheet.set_cell_value(1, 2, Value::number(5.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 1, 0), MakeRange(0, 2, 1, 2)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(5.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(99.0), ctx));
}

TEST(CFEvaluator, UniqueValuesMatchesValueAppearingExactlyOnce) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(7.0));
  harness.sheet.set_cell_value(1, 0, Value::number(7.0));
  harness.sheet.set_cell_value(2, 0, Value::number(9.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::UniqueValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(9.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(7.0), ctx));
}

TEST(CFEvaluator, UniqueValuesNonStorableValueKindDoesNotMatch) {
  // Errors and blanks never match either rule.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(1.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 0, 0)};
  CFRule r = MakeRule(RuleType::UniqueValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
}

TEST(CFEvaluator, DuplicateUniqueValueOnlyOverloadStillReturnsFalse) {
  // Pin the staging contract: value-only overload has no sqref access,
  // so DuplicateValues / UniqueValues continue to return false there.
  CFRule dup = MakeRule(RuleType::DuplicateValues);
  EXPECT_FALSE(match_rule(dup, Value::number(1.0)));
  CFRule uniq = MakeRule(RuleType::UniqueValues);
  EXPECT_FALSE(match_rule(uniq, Value::number(1.0)));
}

// ---------------------------------------------------------------------------
// AboveAverage
// ---------------------------------------------------------------------------
//
// All AboveAverage tests use the population [10, 20, 30, 40, 50] in
// A1:A5 — mean = 30, sample std-dev = sqrt(250) ≈ 15.811.

TEST(CFEvaluator, AboveAverageWithoutSqrefDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  CFRule r = MakeRule(RuleType::AboveAverage);
  // ctx.sqref intentionally nullptr.
  EXPECT_FALSE(match_rule(r, Value::number(40.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, AboveAverageMatchesValuesStrictlyAboveMean) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  r.above_average = true;
  r.equal_average = false;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(40.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));  // mean
  EXPECT_FALSE(match_rule(r, Value::number(20.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(10.0), ctx));
}

TEST(CFEvaluator, AboveAverageEqualAverageIncludesMean) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  r.above_average = true;
  r.equal_average = true;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(40.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(30.0), ctx));  // mean inclusive
  EXPECT_FALSE(match_rule(r, Value::number(29.999), ctx));
}

TEST(CFEvaluator, AboveAverageBelowSideMatchesValuesStrictlyBelowMean) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  r.above_average = false;
  r.equal_average = false;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(20.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(10.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(40.0), ctx));
}

TEST(CFEvaluator, AboveAverageStdDevShiftsThresholdAboveMean) {
  // mean = 30, sample stddev ≈ 15.811. With std_dev = 1, threshold ≈
  // 45.811 → only 50 (above) matches; 40 falls below.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  r.above_average = true;
  r.equal_average = false;
  r.std_dev = 1.0;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(40.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
}

TEST(CFEvaluator, AboveAverageStdDevShiftsThresholdBelowMeanForBelowSide) {
  // mean = 30, sample stddev ≈ 15.811. With std_dev = 1 on the below
  // side, threshold ≈ 14.189 → only 10 matches.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  r.above_average = false;
  r.equal_average = false;
  r.std_dev = 1.0;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(10.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(20.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
}

TEST(CFEvaluator, AboveAverageNonNumericCellDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::text("40"), ctx));
  EXPECT_FALSE(match_rule(r, Value::boolean(true), ctx));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
}

TEST(CFEvaluator, AboveAverageEmptyPopulationDoesNotMatch) {
  CFEvalHarness harness;  // sqref points at cells that are all blank.
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::number(0.0), ctx));
}

TEST(CFEvaluator, AboveAverageBooleansAndTextExcludedFromPopulation) {
  // Population is [10, 50] (numbers only) → mean = 30. Boolean TRUE
  // and text "100" do not contribute.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(10.0));
  harness.sheet.set_cell_value(1, 0, Value::boolean(true));
  harness.sheet.set_cell_value(2, 0, Value::text("100"));
  harness.sheet.set_cell_value(3, 0, Value::number(50.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 3, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(20.0), ctx));
}

TEST(CFEvaluator, AboveAveragePreparedPopulationMatchesSingleCellPathAtStdDevThresholds) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(1.0));
  harness.sheet.set_cell_value(1, 0, Value::number(3.0));
  harness.sheet.set_cell_value(2, 0, Value::number(5.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};

  const ColorScalePopulation population = helpers::gather_population(sqref, harness.sheet);
  ASSERT_EQ(population.sorted.size(), 3U);
  EXPECT_DOUBLE_EQ(population.mean, 3.0);
  EXPECT_DOUBLE_EQ(population.sample_stddev, 2.0);

  const CFEvalContext single_cell_ctx = SqrefContext(harness, sqref);
  CFEvalContext prepared_ctx = single_cell_ctx;
  prepared_ctx.cached_population = &population;

  for (bool equal_average : {false, true}) {
    CFRule above = MakeRule(RuleType::AboveAverage);
    above.above_average = true;
    above.equal_average = equal_average;
    above.std_dev = 1.0;
    for (double value : {4.0, 5.0, 6.0}) {
      const bool single = match_rule(above, Value::number(value), single_cell_ctx);
      const bool prepared = match_rule(above, Value::number(value), prepared_ctx);
      EXPECT_EQ(prepared, single) << "above value=" << value << " equal_average=" << equal_average;
      if (value == 5.0) {
        EXPECT_EQ(prepared, equal_average);
      }
    }

    CFRule below = MakeRule(RuleType::AboveAverage);
    below.above_average = false;
    below.equal_average = equal_average;
    below.std_dev = 1.0;
    for (double value : {0.0, 1.0, 2.0}) {
      const bool single = match_rule(below, Value::number(value), single_cell_ctx);
      const bool prepared = match_rule(below, Value::number(value), prepared_ctx);
      EXPECT_EQ(prepared, single) << "below value=" << value << " equal_average=" << equal_average;
      if (value == 1.0) {
        EXPECT_EQ(prepared, equal_average);
      }
    }
  }
}

TEST(CFEvaluator, PopulationStatisticsUseSortedOrderAndHandleEmptyAndSingleton) {
  CFEvalHarness harness;
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  ColorScalePopulation population = helpers::gather_population(sqref, harness.sheet);
  EXPECT_TRUE(population.sorted.empty());
  EXPECT_DOUBLE_EQ(population.mean, 0.0);
  EXPECT_DOUBLE_EQ(population.sample_stddev, 0.0);

  harness.sheet.set_cell_value(0, 0, Value::number(1.0e16));
  harness.sheet.set_cell_value(1, 0, Value::number(-1.0e16));
  harness.sheet.set_cell_value(2, 0, Value::number(1.0));
  population = helpers::gather_population(sqref, harness.sheet);
  ASSERT_EQ(population.sorted.size(), 3U);
  EXPECT_DOUBLE_EQ(population.sorted[0], -1.0e16);
  EXPECT_DOUBLE_EQ(population.sorted[1], 1.0);
  EXPECT_DOUBLE_EQ(population.sorted[2], 1.0e16);
  // Cancellation-sensitive: the mean differs unless reduced over the sorted values.
  EXPECT_DOUBLE_EQ(population.mean, 0.0);
  EXPECT_DOUBLE_EQ(population.mean, helpers::mean_of(population.sorted));
  EXPECT_DOUBLE_EQ(population.sample_stddev, helpers::sample_stddev(population.sorted, population.mean));

  harness.sheet.set_cell_value(0, 0, Value::number(7.0));
  harness.sheet.set_cell_value(1, 0, Value::blank());
  harness.sheet.set_cell_value(2, 0, Value::blank());
  population = helpers::gather_population(sqref, harness.sheet);
  ASSERT_EQ(population.sorted.size(), 1U);
  EXPECT_DOUBLE_EQ(population.mean, 7.0);
  EXPECT_DOUBLE_EQ(population.sample_stddev, 0.0);
}

TEST(CFEvaluator, OversizedSparsePopulationPreservesSpillsFormulaValuesAndOverlapMultiplicity) {
  CFEvalHarness harness;
  constexpr std::uint32_t kLastRow = 255U;
  constexpr std::uint32_t kLastCol = 256U;
  harness.sheet.set_cell_value(0U, 0U, Value::number(10.0));
  harness.sheet.set_cell_value(kLastRow, kLastCol, Value::number(90.0));
  harness.sheet.set_cell_formula(10U, 10U, "=1+1");
  harness.sheet.set_cell_cached_value(10U, 10U, Value::number(15.0));
  ASSERT_TRUE(harness.sheet.commit_spill(
      128U, 128U, 2U, 2U, {Value::number(20.0), Value::text("20"), Value::boolean(true), Value::number(30.0)}));

  // 65,792 cells takes the oversized-range path while keeping the previous
  // dense workload bounded. The nested sqref deliberately counts each
  // spill value twice, matching the range-union multiplicity contract.
  const std::vector<CFCellRange> sqref{MakeRange(0U, 0U, kLastRow, kLastCol), MakeRange(128U, 128U, 129U, 129U)};
  const std::vector<double> values = helpers::collect_numeric_values(sqref, harness.sheet);
  const std::vector<double> expected{10.0, 15.0, 20.0, 30.0, 90.0, 20.0, 30.0};
  ASSERT_EQ(values.size(), expected.size());
  for (std::size_t i = 0; i < expected.size(); ++i) {
    EXPECT_DOUBLE_EQ(values[i], expected[i]) << "index=" << i;
  }

  EXPECT_EQ(helpers::count_matches_in_sqref(Value::number(20.0), sqref, harness.sheet), 2U);
  const DuplicateValuePopulation duplicates =
      helpers::gather_duplicate_population(sqref, harness.sheet, default_excel_profile());
  EXPECT_EQ(helpers::count_matches_in_population(Value::number(20.0), duplicates), 2U);
  EXPECT_EQ(helpers::count_matches_in_population(Value::text("20"), duplicates), 2U);
  EXPECT_EQ(helpers::count_matches_in_population(Value::boolean(true), duplicates), 2U);
}

TEST(CFEvaluator, AboveAverageValueOnlyOverloadStillReturnsFalse) {
  // Pin the staging contract: AboveAverage continues to return false on
  // the value-only overload (no sqref access).
  CFRule r = MakeRule(RuleType::AboveAverage);
  EXPECT_FALSE(match_rule(r, Value::number(50.0)));
}

// ---------------------------------------------------------------------------
// Top10
// ---------------------------------------------------------------------------

TEST(CFEvaluator, Top10WithoutSqrefDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 2;
  EXPECT_FALSE(match_rule(r, Value::number(50.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, Top10MatchesTopNValuesByRank) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);  // [10, 20, 30, 40, 50]
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 2;
  r.bottom = false;
  r.percent = false;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(40.0), ctx));  // 2nd-largest
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(10.0), ctx));
}

TEST(CFEvaluator, Top10BottomMatchesBottomNValuesByRank) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 2;
  r.bottom = true;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(10.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(20.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(50.0), ctx));
}

TEST(CFEvaluator, Top10TiesAtThresholdAreIncluded) {
  // Population is [40, 40, 40, 30, 10]. Top 2 by rank → threshold = 40,
  // and all three 40s match because the comparison is `>= threshold`.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(40.0));
  harness.sheet.set_cell_value(1, 0, Value::number(40.0));
  harness.sheet.set_cell_value(2, 0, Value::number(40.0));
  harness.sheet.set_cell_value(3, 0, Value::number(30.0));
  harness.sheet.set_cell_value(4, 0, Value::number(10.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 2;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(40.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(10.0), ctx));
}

TEST(CFEvaluator, Top10PercentInterpretsRankAsPercentOfPopulation) {
  // Population size 10 → top 30% = floor(10 * 30 / 100) = 3.
  CFEvalHarness harness;
  for (std::uint32_t row = 0; row < 10; ++row) {
    harness.sheet.set_cell_value(row, 0, Value::number(static_cast<double>(row + 1)));
  }
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 9, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 30;
  r.percent = true;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(10.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(9.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(8.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(7.0), ctx));
}

TEST(CFEvaluator, Top10PercentClampsToAtLeastOne) {
  // Population size 5, percent = 1 → floor(5 * 1 / 100) = 0; clamps to 1.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 1;
  r.percent = true;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(40.0), ctx));
}

TEST(CFEvaluator, Top10RankExceedingPopulationClampsToAll) {
  // Rank 100 with a 5-cell population → threshold = min/max so every
  // cell matches.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 100;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(10.0), ctx));
}

TEST(CFEvaluator, Top10ZeroOrNegativeRankDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 0;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::number(50.0), ctx));
  r.rank = -3;
  const auto ctx2 = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::number(50.0), ctx2));
}

TEST(CFEvaluator, Top10NonNumericCellDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 5;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::text("50"), ctx));
  EXPECT_FALSE(match_rule(r, Value::boolean(true), ctx));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
}

TEST(CFEvaluator, Top10EmptyPopulationDoesNotMatch) {
  CFEvalHarness harness;
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 2;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::number(0.0), ctx));
}

TEST(CFEvaluator, Top10DefaultRankIsTen) {
  // No `rank` set → defaults to 10. With only 5 numbers, that clamps to
  // 5 (the entire population), so every value matches.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  // r.rank intentionally unset.
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(10.0), ctx));
}

TEST(CFEvaluator, Top10ValueOnlyOverloadStillReturnsFalse) {
  // Pin the staging contract: Top10 continues to return false on the
  // value-only overload (no sqref access).
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 5;
  EXPECT_FALSE(match_rule(r, Value::number(50.0)));
}

}  // namespace
}  // namespace formulon::cf
