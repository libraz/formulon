// Conditional-format evaluator tests: range dispatch and population caching.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_evaluator.h"
#include "cf/cf_helpers.h"
#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "cf_evaluator_fixtures.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"
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
// evaluate_cf_at — cross-block priority chain + stopIfTrue
// ---------------------------------------------------------------------------

CFHost MakeHost(CFEvalHarness& harness) {
  CFHost host;
  host.arena = &harness.arena;
  host.registry = &eval::default_registry();
  host.eval_ctx = &harness.eval_ctx;
  return host;
}

ConditionalFormat MakeBlock(std::vector<CFCellRange> sqref, std::vector<CFRule> rules) {
  ConditionalFormat block;
  block.sqref = std::move(sqref);
  block.rules = std::move(rules);
  return block;
}

CFRule MakeBlanksRule(std::int32_t priority, std::uint32_t dxf_id, std::string id, bool stop_if_true = false) {
  CFRule rule;
  rule.type = RuleType::ContainsBlanks;
  rule.priority = priority;
  rule.dxf_id = dxf_id;
  rule.id = std::move(id);
  rule.stop_if_true = stop_if_true;
  return rule;
}

TEST(CFEvaluator, EvaluateCfAtWithoutBlocksReturnsEmpty) {
  CFEvalHarness harness;
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  EXPECT_TRUE(matches.empty());
}

TEST(CFEvaluator, EvaluateCfAtIgnoresBlocksWhoseSqrefExcludesTarget) {
  CFEvalHarness harness;
  // A1 is blank; block sqref is C1:D2 → does not contain A1.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 2, 1, 3)}, {MakeBlanksRule(1, 7, "rule-1")}));
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  EXPECT_TRUE(matches.empty());
}

TEST(CFEvaluator, EvaluateCfAtReturnsMatchingDifferentialFormat) {
  CFEvalHarness harness;
  // A1 is blank by default; ContainsBlanks rule should fire.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 1, 1)}, {MakeBlanksRule(1, 7, "rule-1")}));
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  ASSERT_EQ(matches.size(), 1u);
  EXPECT_EQ(matches[0].rule_id, "rule-1");
  EXPECT_EQ(matches[0].priority, 1);
  EXPECT_EQ(matches[0].kind, CFMatchKind::DifferentialFormat);
  ASSERT_TRUE(matches[0].dxf_id.has_value());
  EXPECT_EQ(*matches[0].dxf_id, 7u);
}

TEST(CFEvaluator, EvaluateCfAtSortsByPriorityAscending) {
  CFEvalHarness harness;
  // Two rules in the same block, declared in reverse priority order.
  // Both fire (cell A1 is blank); evaluator must sort priority=1 first.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 1, 1)},
                {MakeBlanksRule(5, 50, "rule-low-priority"), MakeBlanksRule(1, 10, "rule-high-priority")}));
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  ASSERT_EQ(matches.size(), 2u);
  EXPECT_EQ(matches[0].rule_id, "rule-high-priority");
  EXPECT_EQ(matches[0].priority, 1);
  EXPECT_EQ(matches[1].rule_id, "rule-low-priority");
  EXPECT_EQ(matches[1].priority, 5);
}

TEST(CFEvaluator, EvaluateCfAtSortsAcrossBlocks) {
  CFEvalHarness harness;
  // Two separate blocks — priority is workbook-global so the
  // priority-2 rule from the second block evaluates between the
  // priority-1 and priority-3 rules of the first block.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 1, 1)}, {MakeBlanksRule(1, 11, "p1"), MakeBlanksRule(3, 13, "p3")}));
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {MakeBlanksRule(2, 12, "p2")}));
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  ASSERT_EQ(matches.size(), 3u);
  EXPECT_EQ(matches[0].rule_id, "p1");
  EXPECT_EQ(matches[1].rule_id, "p2");
  EXPECT_EQ(matches[2].rule_id, "p3");
}

TEST(CFEvaluator, EvaluateCfAtFirstWinsFoldSelectsTheHighestPriorityDxf) {
  // Two overlapping rules, neither carrying stopIfTrue, both matching the
  // same cell. The published consumption rule is first-wins over the
  // returned order, so the smaller priority owns the format. Walking the
  // list is what a host does, so the test walks it the documented way.
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 4, 0)}, {MakeBlanksRule(7, 70, "lower-priority")}));
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {MakeBlanksRule(2, 20, "higher-priority")}));

  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  ASSERT_EQ(matches.size(), 2u);
  EXPECT_EQ(matches[0].rule_id, "higher-priority");

  std::optional<std::uint32_t> painted;
  for (const CFMatch& match : matches) {
    if (!painted.has_value() && match.dxf_id.has_value()) {
      painted = match.dxf_id;
    }
  }
  ASSERT_TRUE(painted.has_value());
  EXPECT_EQ(*painted, 20u);
}

TEST(CFEvaluator, EvaluateCfAtStopIfTrueHaltsEvaluation) {
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 1, 1)},
                {MakeBlanksRule(1, 10, "first", /*stop_if_true=*/true), MakeBlanksRule(2, 20, "second")}));
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  ASSERT_EQ(matches.size(), 1u);
  EXPECT_EQ(matches[0].rule_id, "first");
}

TEST(CFEvaluator, EvaluateCfAtStopIfTrueDoesNotHaltOnNonMatch) {
  // First rule does not match (cell is non-blank), so stop_if_true is
  // never triggered; the second rule still evaluates.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(7.0));

  CFRule cell_is_rule;
  cell_is_rule.type = RuleType::CellIs;
  cell_is_rule.priority = 1;
  cell_is_rule.dxf_id = 100;
  cell_is_rule.id = "cell-is-zero";
  cell_is_rule.op = CellIsOperator::Equal;
  cell_is_rule.formula1 = "0";
  cell_is_rule.stop_if_true = true;

  CFRule errors_rule;
  errors_rule.type = RuleType::NotContainsErrors;
  errors_rule.priority = 2;
  errors_rule.dxf_id = 200;
  errors_rule.id = "not-error";

  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {cell_is_rule, errors_rule}));
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  ASSERT_EQ(matches.size(), 1u);
  EXPECT_EQ(matches[0].rule_id, "not-error");
}

TEST(CFEvaluator, EvaluateCfAtVisualRulePopulatesRenderPayload) {
  // ColorScale rule on A1:A5; cell A3 is the population midpoint.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  CFRule color_rule = MakeRule(RuleType::ColorScale);
  color_rule.priority = 1;
  color_rule.id = "color-1";
  color_rule.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  harness.sheet.mutable_conditional_formats().push_back(MakeBlock({MakeRange(0, 0, 4, 0)}, {color_rule}));

  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(2, 0), host);
  ASSERT_EQ(matches.size(), 1u);
  EXPECT_EQ(matches[0].kind, CFMatchKind::ColorScale);
  ASSERT_TRUE(matches[0].resolved_fill_color.has_value());
  EXPECT_EQ(matches[0].resolved_fill_color->r, 128);
  EXPECT_EQ(matches[0].resolved_fill_color->g, 128);
  EXPECT_EQ(matches[0].resolved_fill_color->b, 0);
}

TEST(CFEvaluator, EvaluateCfAtMissingHostFieldsReturnsEmpty) {
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {MakeBlanksRule(1, 7, "rule-1")}));
  CFHost host;  // arena/registry/eval_ctx all null.
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  EXPECT_TRUE(matches.empty());
}

// ---------------------------------------------------------------------------
// evaluate_cf_for_range — viewport-range API
// ---------------------------------------------------------------------------

TEST(CFEvaluator, EvaluateCfForRangeReturnsOneEntryPerMatchedCell) {
  CFEvalHarness harness;
  // A1 and A2 are blank → both match ContainsBlanks. A3 has a value.
  harness.sheet.set_cell_value(2, 0, Value::number(42.0));
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 2, 0)}, {MakeBlanksRule(1, 7, "blanks")}));

  const auto host = MakeHost(harness);
  std::vector<CFRangeCellMatches> matches = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 2, 0), host).value();
  ASSERT_EQ(matches.size(), 2u);
  EXPECT_EQ(matches[0].cell.row, 0u);
  EXPECT_EQ(matches[0].cell.col, 0u);
  EXPECT_EQ(matches[1].cell.row, 1u);
  EXPECT_EQ(matches[1].cell.col, 0u);
  ASSERT_EQ(matches[0].matches.size(), 1u);
  EXPECT_EQ(matches[0].matches[0].rule_id, "blanks");
}

TEST(CFEvaluator, EvaluateCfForRangeEmitsRowMajorOrder) {
  CFEvalHarness harness;
  // 2x2 viewport, all blank → all 4 cells match.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 1, 1)}, {MakeBlanksRule(1, 7, "blanks")}));

  const auto host = MakeHost(harness);
  std::vector<CFRangeCellMatches> matches = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 1, 1), host).value();
  ASSERT_EQ(matches.size(), 4u);
  // Row-major: (0,0), (0,1), (1,0), (1,1).
  EXPECT_EQ(matches[0].cell.row, 0u);
  EXPECT_EQ(matches[0].cell.col, 0u);
  EXPECT_EQ(matches[1].cell.row, 0u);
  EXPECT_EQ(matches[1].cell.col, 1u);
  EXPECT_EQ(matches[2].cell.row, 1u);
  EXPECT_EQ(matches[2].cell.col, 0u);
  EXPECT_EQ(matches[3].cell.row, 1u);
  EXPECT_EQ(matches[3].cell.col, 1u);
}

TEST(CFEvaluator, EvaluateCfForRangeSkipsCellsWithNoMatches) {
  CFEvalHarness harness;
  // Block applies only to A1; B1 has no rules.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {MakeBlanksRule(1, 7, "blanks")}));
  const auto host = MakeHost(harness);
  // Viewport spans A1:B1 → only A1 should appear in the result.
  std::vector<CFRangeCellMatches> matches = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 0, 1), host).value();
  ASSERT_EQ(matches.size(), 1u);
  EXPECT_EQ(matches[0].cell.row, 0u);
  EXPECT_EQ(matches[0].cell.col, 0u);
}

TEST(CFEvaluator, EvaluateCfForRangeSingleCellRangeStillVisited) {
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {MakeBlanksRule(1, 7, "blanks")}));
  const auto host = MakeHost(harness);
  // first == last (A1:A1).
  std::vector<CFRangeCellMatches> matches = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 0, 0), host).value();
  ASSERT_EQ(matches.size(), 1u);
  EXPECT_EQ(matches[0].cell.row, 0u);
  EXPECT_EQ(matches[0].cell.col, 0u);
}

TEST(CFEvaluator, EvaluateCfForRangeReturnsEmptyForEmptyHost) {
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {MakeBlanksRule(1, 7, "blanks")}));
  CFHost host;  // null fields.
  std::vector<CFRangeCellMatches> matches = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 0, 0), host).value();
  EXPECT_TRUE(matches.empty());
}

TEST(CFEvaluator, EvaluateCfForRangeAggregatesPriorityOrderPerCell) {
  // Two rules at different priorities; both match every cell. Pin that
  // each cell's match list comes back in priority order.
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 1)}, {MakeBlanksRule(2, 20, "later"), MakeBlanksRule(1, 10, "earlier")}));
  const auto host = MakeHost(harness);
  std::vector<CFRangeCellMatches> matches = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 0, 1), host).value();
  ASSERT_EQ(matches.size(), 2u);
  for (const auto& cell : matches) {
    ASSERT_EQ(cell.matches.size(), 2u);
    EXPECT_EQ(cell.matches[0].rule_id, "earlier");
    EXPECT_EQ(cell.matches[1].rule_id, "later");
  }
}

CFRule MakeGreaterThanRule(std::int32_t priority, std::uint32_t dxf_id, std::string id, std::string threshold) {
  CFRule rule;
  rule.type = RuleType::CellIs;
  rule.op = CellIsOperator::GreaterThan;
  rule.priority = priority;
  rule.dxf_id = dxf_id;
  rule.id = std::move(id);
  rule.formula1 = std::move(threshold);
  return rule;
}

TEST(CFEvaluator, EvaluateCfForRangeWholeColumnWalksOnlyTheBlock) {
  // One 3x3 block on a sheet; the request is the whole of column A. The
  // walk must cost the block's overlap with the column (three cells),
  // not the column's million rows — this test finishing inside the fast
  // tier is that bound's evidence.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(10.0));
  harness.sheet.set_cell_value(1, 0, Value::number(60.0));
  harness.sheet.set_cell_value(2, 0, Value::number(90.0));
  harness.sheet.set_cell_value(1, 1, Value::number(70.0));
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 2, 2)}, {MakeGreaterThanRule(1, 7, "gt50", "50")}));

  const auto host = MakeHost(harness);
  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, kCfMaxRows - 1U, 0), host);
  ASSERT_TRUE(static_cast<bool>(result));
  const std::vector<CFRangeCellMatches>& matches = result.value();
  ASSERT_EQ(matches.size(), 2u);
  EXPECT_EQ(matches[0].cell.row, 1u);
  EXPECT_EQ(matches[0].cell.col, 0u);
  ASSERT_EQ(matches[0].matches.size(), 1u);
  EXPECT_EQ(matches[0].matches[0].rule_id, "gt50");
  EXPECT_EQ(matches[0].matches[0].dxf_id, 7u);
  EXPECT_EQ(matches[1].cell.row, 2u);
  EXPECT_EQ(matches[1].cell.col, 0u);
  ASSERT_EQ(matches[1].matches.size(), 1u);
  EXPECT_EQ(matches[1].matches[0].rule_id, "gt50");
}

TEST(CFEvaluator, EvaluateCfForRangeRejectsRequestPastViewportCeiling) {
  // Two full columns are past the viewport ceiling: the request is
  // refused before any cell is visited.
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {MakeBlanksRule(1, 7, "blanks")}));

  const auto host = MakeHost(harness);
  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, kCfMaxRows - 1U, 1), host);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kSecResourceLimit);
}

TEST(CFEvaluator, EvaluateCfForRangeOverlappingBlocksVisitEachCellOnce) {
  // Two blocks share column B. Every cell of the viewport must appear
  // exactly once and in row-major order, with the shared cells carrying
  // both blocks' matches in priority order.
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 1, 1)}, {MakeBlanksRule(2, 20, "left")}));
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 1, 1, 2)}, {MakeBlanksRule(1, 10, "right")}));

  const auto host = MakeHost(harness);
  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 1, 2), host);
  ASSERT_TRUE(static_cast<bool>(result));
  const std::vector<CFRangeCellMatches>& matches = result.value();
  ASSERT_EQ(matches.size(), 6u);
  const std::vector<std::pair<std::uint32_t, std::uint32_t>> expected_cells{{0, 0}, {0, 1}, {0, 2},
                                                                            {1, 0}, {1, 1}, {1, 2}};
  for (std::size_t i = 0; i < expected_cells.size(); ++i) {
    EXPECT_EQ(matches[i].cell.row, expected_cells[i].first) << "i=" << i;
    EXPECT_EQ(matches[i].cell.col, expected_cells[i].second) << "i=" << i;
  }
  // Column A sees only the left block, column C only the right one, and
  // column B both — priority-ascending.
  ASSERT_EQ(matches[0].matches.size(), 1u);
  EXPECT_EQ(matches[0].matches[0].rule_id, "left");
  ASSERT_EQ(matches[1].matches.size(), 2u);
  EXPECT_EQ(matches[1].matches[0].rule_id, "right");
  EXPECT_EQ(matches[1].matches[1].rule_id, "left");
  ASSERT_EQ(matches[2].matches.size(), 1u);
  EXPECT_EQ(matches[2].matches[0].rule_id, "right");
}

TEST(CFEvaluator, EvaluateCfForRangeClipsBlocksOutsideTheRequest) {
  // A block sitting entirely outside the viewport contributes nothing,
  // and a block straddling the viewport edge contributes only the cells
  // inside it.
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 3, 3)}, {MakeBlanksRule(1, 7, "straddling")}));
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(10, 10, 12, 12)}, {MakeBlanksRule(1, 8, "far-away")}));

  const auto host = MakeHost(harness);
  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(1, 1, 2, 2), host);
  ASSERT_TRUE(static_cast<bool>(result));
  const std::vector<CFRangeCellMatches>& matches = result.value();
  ASSERT_EQ(matches.size(), 4u);
  EXPECT_EQ(matches[0].cell.row, 1u);
  EXPECT_EQ(matches[0].cell.col, 1u);
  EXPECT_EQ(matches[3].cell.row, 2u);
  EXPECT_EQ(matches[3].cell.col, 2u);
  for (const auto& cell : matches) {
    ASSERT_EQ(cell.matches.size(), 1u);
    EXPECT_EQ(cell.matches[0].rule_id, "straddling");
  }
}

// ---------------------------------------------------------------------------
// Population caching — viewport API must produce results identical to
// per-cell `evaluate_cf_at` calls. This pins the optimisation: caching
// the population across cells in a range may not change rendered output.
// ---------------------------------------------------------------------------

CFRule MakeColorScaleBlock(std::int32_t priority, std::string id) {
  CFRule rule;
  rule.type = RuleType::ColorScale;
  rule.priority = priority;
  rule.id = std::move(id);
  rule.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  return rule;
}

CFRule MakeTop10Rule(std::int32_t priority, std::string id, std::int32_t rank) {
  CFRule rule;
  rule.type = RuleType::Top10;
  rule.priority = priority;
  rule.dxf_id = 5;
  rule.id = std::move(id);
  rule.rank = rank;
  rule.bottom = false;
  rule.percent = false;
  return rule;
}

TEST(CFEvaluator, EvaluateCfForRangeCachedPathMatchesUncached) {
  // ColorScale over A1:A5 = [10, 20, 30, 40, 50]. Every cell is in
  // the sqref so `evaluate_cf_at` and `evaluate_cf_for_range` must
  // resolve identical fill colours.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 4, 0)}, {MakeColorScaleBlock(1, "color-scale")}));

  const auto host = MakeHost(harness);
  std::vector<CFRangeCellMatches> ranged = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 4, 0), host).value();
  ASSERT_EQ(ranged.size(), 5u);

  for (std::uint32_t row = 0; row < 5; ++row) {
    std::vector<CFMatch> per_cell = evaluate_cf_at(harness.sheet, At(row, 0), host);
    ASSERT_EQ(per_cell.size(), 1u) << "row=" << row;
    ASSERT_EQ(ranged[row].matches.size(), 1u) << "row=" << row;
    const auto& cached = ranged[row].matches[0];
    const auto& fresh = per_cell[0];
    ASSERT_TRUE(cached.resolved_fill_color.has_value()) << "row=" << row;
    ASSERT_TRUE(fresh.resolved_fill_color.has_value()) << "row=" << row;
    EXPECT_EQ(cached.resolved_fill_color->r, fresh.resolved_fill_color->r) << "row=" << row;
    EXPECT_EQ(cached.resolved_fill_color->g, fresh.resolved_fill_color->g) << "row=" << row;
    EXPECT_EQ(cached.resolved_fill_color->b, fresh.resolved_fill_color->b) << "row=" << row;
    EXPECT_EQ(cached.resolved_fill_color->a, fresh.resolved_fill_color->a) << "row=" << row;
  }
}

TEST(CFEvaluator, EvaluateCfForRangeCachedPathPreservesTop10Behavior) {
  // Top-2 over [10, 20, 30, 40, 50] should match rows 3 (40) and 4
  // (50). Cached and uncached paths must agree on which cells fire.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 4, 0)}, {MakeTop10Rule(1, "top2", 2)}));

  const auto host = MakeHost(harness);
  std::vector<CFRangeCellMatches> ranged = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 4, 0), host).value();

  // Row-major: only rows 3 and 4 should appear.
  ASSERT_EQ(ranged.size(), 2u);
  EXPECT_EQ(ranged[0].cell.row, 3u);
  EXPECT_EQ(ranged[1].cell.row, 4u);

  // Cross-check by walking cell-by-cell with the uncached entry point.
  for (std::uint32_t row = 0; row < 5; ++row) {
    std::vector<CFMatch> per_cell = evaluate_cf_at(harness.sheet, At(row, 0), host);
    const bool expected = row >= 3;
    EXPECT_EQ(per_cell.size(), expected ? 1u : 0u) << "row=" << row;
  }
}

// A rule whose sqref is an explicit multi-million-cell rectangle (not
// whole-column / whole-row notation) must be clamped to the populated
// extent before scanning, so the scan stays bounded and still returns the
// same match/numeric results as the tight rectangle would.
TEST(CFHelpers, ExplicitGiantSqrefIsClampedToPopulatedExtent) {
  Sheet sheet("S");
  sheet.set_cell_cached_value(0, 0, Value::number(5.0));
  sheet.set_cell_cached_value(50, 5, Value::number(5.0));
  sheet.set_cell_cached_value(100, 3, Value::number(5.0));
  sheet.set_cell_cached_value(100, 10, Value::number(9.0));  // populated, non-matching

  // Explicit rectangle far larger than the ~65k clamp threshold, yet not
  // full-column (last.row != kCfMaxRows-1) nor full-row (last.col !=
  // kCfMaxCols-1) — the case the old code walked cell-by-cell.
  CFCellRange giant;
  giant.first = CellAddress{0, 0};
  giant.last = CellAddress{400000, 4000};
  ASSERT_FALSE(giant.is_full_col());
  ASSERT_FALSE(giant.is_full_row());
  const std::vector<CFCellRange> sqref{giant};

  EXPECT_EQ(helpers::count_matches_in_sqref(Value::number(5.0), sqref, sheet), 3u);

  const std::vector<double> numbers = helpers::collect_numeric_values(sqref, sheet);
  EXPECT_EQ(numbers.size(), 4u);  // three 5s + one 9
  double sum = 0.0;
  for (double n : numbers) {
    sum += n;
  }
  EXPECT_DOUBLE_EQ(sum, 24.0);
}

TEST(CFHelpers, WholeAxisExtentIncludesPhantomOnlyColumn) {
  Sheet sheet("S");
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 2U,
                                 {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(4.0),
                                  Value::number(5.0), Value::number(6.0)}));

  const std::vector<CFCellRange> sqref{MakeRange(0U, 1U, kCfMaxRows - 1U, 1U)};
  EXPECT_EQ(helpers::count_matches_in_sqref(Value::number(2.0), sqref, sheet), 1U);
  EXPECT_EQ(helpers::count_matches_in_sqref(Value::number(4.0), sqref, sheet), 1U);
  const std::vector<double> values = helpers::collect_numeric_values(sqref, sheet);
  ASSERT_EQ(values.size(), 3U);
  EXPECT_DOUBLE_EQ(values[0], 2.0);
  EXPECT_DOUBLE_EQ(values[1], 4.0);
  EXPECT_DOUBLE_EQ(values[2], 6.0);
}

}  // namespace
}  // namespace formulon::cf
