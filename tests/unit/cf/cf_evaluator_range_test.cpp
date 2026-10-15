// Conditional-format evaluator tests: range dispatch and population caching.

#include <algorithm>
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

std::uint32_t g_cf_formula_calls = 0;

Value CfFormulaCallCounter(const Value*, std::uint32_t, Arena&) {
  ++g_cf_formula_calls;
  return Value::number(static_cast<double>(g_cf_formula_calls));
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

TEST(CFEvaluator, EvaluateCfAtEqualPriorityStopIfTrueKeepsDocumentOrder) {
  CFEvalHarness harness;
  // Equal priorities are legal in the in-memory model; the first rule in
  // workbook order owns the stopIfTrue decision.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)},
                {MakeBlanksRule(1, 10, "first", /*stop_if_true=*/true), MakeBlanksRule(1, 20, "second")}));
  const auto host = MakeHost(harness);
  const std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
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

TEST(CFEvaluator, PreparedRelativeFormulaReusesScratchWithinSmallCallerArenaBudget) {
  constexpr std::uint32_t kCellCount = 2000U;
  constexpr std::size_t kArenaLimit = 16U * 1024U;
  CFEvalHarness harness;
  for (std::uint32_t row = 0; row < kCellCount; ++row) {
    harness.sheet.set_cell_value(row, 0U, Value::number(static_cast<double>(row + 1U)));
  }

  CFRule rule = MakeRule(RuleType::Expression);
  rule.priority = 1;
  rule.id = "relative-positive";
  rule.formula1 = "A1>0";
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0U, 0U, kCellCount - 1U, 0U)}, {std::move(rule)}));

  Arena arena(512U, kArenaLimit);
  std::uint32_t* const sentinel = arena.create<std::uint32_t>(0xC0FFEEU);
  ASSERT_NE(nullptr, sentinel);
  const std::size_t used_before = arena.bytes_used();
  const std::size_t allocated_before = arena.bytes_allocated();
  const std::size_t chunks_before = arena.chunk_count();
  const bool exhausted_before = arena.exhausted();

  CFHost host = MakeHost(harness);
  host.arena = &arena;
  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0U, 0U, kCellCount - 1U, 0U), host);

  ASSERT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  ASSERT_EQ(result.value().size(), kCellCount);
  for (std::uint32_t row = 0; row < kCellCount; ++row) {
    EXPECT_EQ(result.value()[row].cell, At(row, 0U)) << "row=" << row;
    ASSERT_EQ(result.value()[row].matches.size(), 1U) << "row=" << row;
    EXPECT_EQ(result.value()[row].matches[0].rule_id, "relative-positive") << "row=" << row;
  }

  EXPECT_EQ(host.arena, &arena);
  EXPECT_EQ(*sentinel, 0xC0FFEEU);
  EXPECT_EQ(arena.bytes_used(), used_before);
  EXPECT_EQ(arena.bytes_allocated(), allocated_before);
  EXPECT_EQ(arena.chunk_count(), chunks_before);
  EXPECT_EQ(arena.exhausted(), exhausted_before);
}

TEST(CFEvaluator, RangeEvaluatorReportsParseExhaustionWithoutPartialResults) {
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0U, 0U, 0U, 0U)}, {MakeBlanksRule(1, 7, "already-matched")}));

  CFRule expression = MakeRule(RuleType::Expression);
  expression.priority = 2;
  expression.id = "parse-exhaustion";
  expression.formula1 = "A1>0";
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(1U, 0U, 1U, 0U)}, {std::move(expression)}));

  Arena arena(512U, 0U);
  CFHost host = MakeHost(harness);
  host.arena = &arena;
  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0U, 0U, 1U, 0U), host);

  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kSecResourceLimit);
  EXPECT_NE(result.error().context.find("cell="), std::string::npos);
  EXPECT_NE(result.error().context.find("ceiling="), std::string::npos);
}

TEST(CFEvaluator, ParseShiftEvaluationLatchesScratchExhaustionForNonzeroShift) {
  CFEvalHarness harness;
  Arena arena(4096U, 4096U);
  PreparedCfFormulaCache cache;
  CFEvalContext ctx = harness.context(At(0U, 0U), At(1U, 0U));
  ctx.arena = &arena;
  ctx.formula_cache = &cache;

  const Value result = helpers::parse_shift_evaluate("A1>100", ctx);

  EXPECT_TRUE(result.is_error());
  ASSERT_TRUE(cache.persistent_arena.has_value());
  ASSERT_EQ(cache.roots.size(), 1U);
  EXPECT_NE(nullptr, cache.roots[0]);
  EXPECT_EQ(cache.persistent_arena->max_bytes(), 4096U);
  EXPECT_EQ(cache.persistent_arena->bytes_allocated(), 4096U);
  ASSERT_TRUE(cache.scratch_arena.has_value());
  EXPECT_EQ(cache.scratch_arena->max_bytes(), 0U);
  EXPECT_TRUE(cache.scratch_arena->exhausted());
  EXPECT_TRUE(cache.resource_failure);
}

TEST(CFEvaluator, RangeEvaluatorReportsShiftExhaustionEvenWhenRuleDoesNotMatch) {
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0U, 0U, 0U, 0U)}, {MakeBlanksRule(1, 7, "already-matched", true)}));

  CFRule expression = MakeRule(RuleType::Expression);
  expression.priority = 2;
  expression.id = "shift-exhaustion";
  expression.formula1 = "ROW(A1)=1";  // At A2 this shifts to ROW(A2)=1, which is false.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0U, 0U, 1U, 0U)}, {std::move(expression)}));

  // Prove A2 reaches the expression and the formula shifts when scratch space
  // is available: the unshifted formula would match at A2. A1's earlier
  // stopIfTrue rule avoids parsing it at the anchor.
  Arena control_arena;
  CFHost control_host = MakeHost(harness);
  control_host.arena = &control_arena;
  const auto control = evaluate_cf_for_range(harness.sheet, MakeRange(0U, 0U, 1U, 0U), control_host);
  ASSERT_TRUE(static_cast<bool>(control)) << (control ? "" : control.error().message);
  ASSERT_EQ(control.value().size(), 1U);
  EXPECT_EQ(control.value()[0].cell, At(0U, 0U));
  ASSERT_EQ(control.value()[0].matches.size(), 1U);
  EXPECT_EQ(control.value()[0].matches[0].rule_id, "already-matched");

  // At A2 the parsed root needs a shifted copy, but its persistent root
  // already uses the entire caller-sized chunk so scratch has no capacity.
  Arena arena(4096U, 4096U);
  CFHost host = MakeHost(harness);
  host.arena = &arena;
  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0U, 0U, 1U, 0U), host);

  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kSecResourceLimit);
  EXPECT_NE(result.error().context.find("cell="), std::string::npos);
  EXPECT_NE(result.error().context.find("ceiling="), std::string::npos);
}

TEST(CFEvaluator, RecursiveComputedTextUsesPrivateMemoAcrossTargetsAndAnchors) {
  CFEvalHarness harness;
  harness.sheet.set_cell_formula(0U, 0U, "=\"match\"");  // A1
  harness.sheet.set_cell_formula(1U, 0U, "=\"no\"");     // A2
  harness.sheet.set_cell_formula(0U, 1U, "=\"no\"");     // B1
  harness.sheet.set_cell_formula(1U, 1U, "=\"match\"");  // B2

  const auto make_expression = [](std::int32_t priority, std::string id) {
    CFRule rule = MakeRule(RuleType::Expression);
    rule.priority = priority;
    rule.id = std::move(id);
    rule.formula1 = "A1=\"match\"";
    return rule;
  };
  // The same source root is reused by two rules. At C1:C2 it reads A1:A2;
  // at B1:C2 it reads the cells one column to the right of those references.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0U, 2U, 1U, 2U)}, {make_expression(1, "first-anchor")}));
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0U, 1U, 1U, 2U)}, {make_expression(2, "second-anchor")}));

  eval::EvalState caller_state;
  caller_state.memoize(&harness.sheet, 9U, 9U, Value::text("caller-owned"));
  const Value* const caller_memo = caller_state.lookup_memo(&harness.sheet, 9U, 9U);
  ASSERT_NE(nullptr, caller_memo);
  eval::EvalContext eval_ctx(harness.sheet, caller_state);
  CFHost host = MakeHost(harness);
  host.eval_ctx = &eval_ctx;

  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0U, 1U, 1U, 2U), host);

  ASSERT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  ASSERT_EQ(result.value().size(), 3U);
  EXPECT_EQ(result.value()[0].cell, At(0U, 1U));
  ASSERT_EQ(result.value()[0].matches.size(), 1U);
  EXPECT_EQ(result.value()[0].matches[0].rule_id, "second-anchor");
  EXPECT_EQ(result.value()[1].cell, At(0U, 2U));
  ASSERT_EQ(result.value()[1].matches.size(), 1U);
  EXPECT_EQ(result.value()[1].matches[0].rule_id, "first-anchor");
  EXPECT_EQ(result.value()[2].cell, At(1U, 2U));
  ASSERT_EQ(result.value()[2].matches.size(), 1U);
  EXPECT_EQ(result.value()[2].matches[0].rule_id, "second-anchor");

  EXPECT_EQ(caller_state.lookup_memo(&harness.sheet, 9U, 9U), caller_memo);
  EXPECT_EQ(caller_state.lookup_memo(&harness.sheet, 0U, 0U), nullptr);
  EXPECT_EQ(caller_state.lookup_memo(&harness.sheet, 1U, 0U), nullptr);
  EXPECT_EQ(caller_state.lookup_memo(&harness.sheet, 0U, 1U), nullptr);
  EXPECT_EQ(caller_state.lookup_memo(&harness.sheet, 1U, 1U), nullptr);
  EXPECT_EQ(caller_state.depth(), 0U);
  EXPECT_FALSE(caller_state.out_of_memory());
}

TEST(CFEvaluator, PreparedFormulaPreservesReadObserverCachedValuesAndDependencies) {
  CFEvalHarness harness;
  harness.sheet.set_cell_formula(0U, 0U, "=1");
  harness.sheet.set_cell_cached_value(0U, 0U, Value::number(0.0));

  CFRule expression = MakeRule(RuleType::Expression);
  expression.priority = 1;
  expression.id = "observer-cached-value";
  expression.formula1 = "A1>0";
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0U, 1U, 0U, 1U)}, {std::move(expression)}));

  eval::EvalState observer;
  observer.observe_reads_of(&harness.sheet, 0U, 0U, 42U);
  observer.memoize(&harness.sheet, 9U, 9U, Value::text("caller-owned"));
  const Value* const caller_memo = observer.lookup_memo(&harness.sheet, 9U, 9U);
  ASSERT_NE(nullptr, caller_memo);
  eval::EvalContext eval_ctx(harness.sheet, observer);
  CFHost host = MakeHost(harness);
  host.eval_ctx = &eval_ctx;

  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0U, 1U, 0U, 1U), host);

  ASSERT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  EXPECT_TRUE(result.value().empty());  // A1's cached 0 is observed, not its formula text (=1).
  ASSERT_EQ(observer.observed_reads().size(), 1U);
  EXPECT_EQ(observer.observed_reads()[0], 42U);
  EXPECT_EQ(observer.lookup_memo(&harness.sheet, 9U, 9U), caller_memo);
  EXPECT_EQ(observer.lookup_memo(&harness.sheet, 0U, 0U), nullptr);
  EXPECT_EQ(observer.depth(), 0U);
}

TEST(CFEvaluator, RangeCellIsBetweenRetainsBothFormulaOperands) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0U, 0U, Value::text("A"));
  harness.sheet.set_cell_value(0U, 1U, Value::text("Z"));
  harness.sheet.set_cell_value(0U, 2U, Value::text("M"));

  CFRule between = MakeRule(RuleType::CellIs);
  between.op = CellIsOperator::Between;
  between.formula1 = "A1&\"\"";
  between.formula2 = "B1&\"\"";
  harness.sheet.mutable_conditional_formats().push_back(MakeBlock({MakeRange(0U, 2U, 0U, 2U)}, {std::move(between)}));

  const auto host = MakeHost(harness);
  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0U, 2U, 0U, 2U), host);

  ASSERT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  ASSERT_EQ(result.value().size(), 1U);
  ASSERT_EQ(result.value()[0].matches.size(), 1U);
  EXPECT_EQ(result.value()[0].matches[0].kind, CFMatchKind::DifferentialFormat);
}

TEST(CFEvaluator, RangeColorScaleEvaluatesTwoFormulaCfvosFromTheirAnchor) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0U, 0U, Value::number(5.0));
  harness.sheet.set_cell_value(0U, 1U, Value::number(5.0));
  harness.sheet.set_cell_value(1U, 1U, Value::number(10.0));
  harness.sheet.set_cell_value(2U, 1U, Value::number(15.0));

  CFRule scale = MakeRule(RuleType::ColorScale);
  scale.priority = 1;
  scale.id = "formula-cfvos";
  scale.color_scale = TwoStopMinMax(RGB(255U, 0U, 0U), RGB(0U, 255U, 0U));
  scale.color_scale->thresholds = {test::Cfvo(CfvoType::Formula, "A1"), test::Cfvo(CfvoType::Formula, "A1+10")};
  harness.sheet.mutable_conditional_formats().push_back(MakeBlock({MakeRange(0U, 1U, 2U, 1U)}, {std::move(scale)}));

  const auto host = MakeHost(harness);
  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0U, 1U, 2U, 1U), host);

  ASSERT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  ASSERT_EQ(result.value().size(), 3U);
  ASSERT_TRUE(result.value().front().matches[0].resolved_fill_color.has_value());
  ASSERT_TRUE(result.value().back().matches[0].resolved_fill_color.has_value());
  EXPECT_EQ(*result.value().front().matches[0].resolved_fill_color, RGB(255U, 0U, 0U));
  EXPECT_EQ(*result.value().back().matches[0].resolved_fill_color, RGB(0U, 255U, 0U));
}

TEST(CFEvaluator, StopIfTrueSkipsAnUnreachableLargeFormula) {
  CFEvalHarness harness;
  CFRule stop = MakeBlanksRule(1, 7, "stop", true);
  CFRule unreachable = MakeRule(RuleType::Expression);
  unreachable.priority = 2;
  unreachable.id = "unreachable";
  std::string large_formula = "A1";
  for (std::size_t i = 0; i < 1024U; ++i) {
    large_formula += "+A1";
  }
  unreachable.formula1 = std::move(large_formula);
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0U, 0U, 0U, 0U)}, {std::move(stop), std::move(unreachable)}));

  Arena arena(128U, 0U);
  CFHost host = MakeHost(harness);
  host.arena = &arena;
  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0U, 0U, 0U, 0U), host);

  ASSERT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  ASSERT_EQ(result.value().size(), 1U);
  ASSERT_EQ(result.value()[0].matches.size(), 1U);
  EXPECT_EQ(result.value()[0].matches[0].rule_id, "stop");
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

TEST(CFEvaluator, EvaluateCfForRangeCachedAboveAverageMatchesSingleCellEvaluation) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0U, 0U, Value::number(1.0));
  harness.sheet.set_cell_value(1U, 0U, Value::number(3.0));
  harness.sheet.set_cell_value(2U, 0U, Value::number(5.0));

  const auto make_above_average = [](std::int32_t priority, std::string id, bool above_average, bool equal_average) {
    CFRule rule = MakeRule(RuleType::AboveAverage);
    rule.priority = priority;
    rule.id = std::move(id);
    rule.above_average = above_average;
    rule.equal_average = equal_average;
    rule.std_dev = 1.0;
    return rule;
  };
  harness.sheet.mutable_conditional_formats().push_back(MakeBlock(
      {MakeRange(0U, 0U, 2U, 0U)},
      {make_above_average(1, "above-strict", true, false), make_above_average(2, "above-inclusive", true, true),
       make_above_average(3, "below-strict", false, false), make_above_average(4, "below-inclusive", false, true)}));

  const auto host = MakeHost(harness);
  const auto ranged = evaluate_cf_for_range(harness.sheet, MakeRange(0U, 0U, 2U, 0U), host);
  ASSERT_TRUE(static_cast<bool>(ranged));
  ASSERT_EQ(ranged.value().size(), 2U);
  EXPECT_EQ(ranged.value()[0].cell, At(0U, 0U));
  ASSERT_EQ(ranged.value()[0].matches.size(), 1U);
  EXPECT_EQ(ranged.value()[0].matches[0].rule_id, "below-inclusive");
  EXPECT_EQ(ranged.value()[1].cell, At(2U, 0U));
  ASSERT_EQ(ranged.value()[1].matches.size(), 1U);
  EXPECT_EQ(ranged.value()[1].matches[0].rule_id, "above-inclusive");

  for (std::uint32_t row = 0U; row < 3U; ++row) {
    const std::vector<CFMatch> single = evaluate_cf_at(harness.sheet, At(row, 0U), host);
    const auto ranged_cell = std::find_if(ranged.value().begin(), ranged.value().end(),
                                          [row](const CFRangeCellMatches& cell) { return cell.cell.row == row; });
    const std::size_t ranged_count = ranged_cell == ranged.value().end() ? 0U : ranged_cell->matches.size();
    ASSERT_EQ(ranged_count, single.size()) << "row=" << row;
    if (ranged_count != 0U) {
      EXPECT_EQ(ranged_cell->matches[0].rule_id, single[0].rule_id) << "row=" << row;
    }
  }
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

TEST(CFEvaluator, EvaluateCfForRangeDuplicateCachePreservesOverlappingSqrefMultiplicity) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(7.0));
  CFRule duplicate = MakeRule(RuleType::DuplicateValues);
  duplicate.priority = 1;
  duplicate.id = "duplicate";
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0), MakeRange(0, 0, 0, 0)}, {duplicate}));

  const auto host = MakeHost(harness);
  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 0, 0), host);
  ASSERT_TRUE(static_cast<bool>(result));
  ASSERT_EQ(result.value().size(), 1u);
  ASSERT_EQ(result.value()[0].matches.size(), 1u);
  EXPECT_EQ(result.value()[0].matches[0].rule_id, "duplicate");
}

TEST(CFEvaluator, EvaluateCfForRangeDuplicateCacheFoldsSignedZeroAndKeepsKindsApart) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(0.0));
  harness.sheet.set_cell_value(1, 0, Value::number(-0.0));
  harness.sheet.set_cell_value(2, 0, Value::number(1.0));
  harness.sheet.set_cell_value(3, 0, Value::boolean(true));
  harness.sheet.set_cell_value(4, 0, Value::text("1"));
  CFRule duplicate = MakeRule(RuleType::DuplicateValues);
  duplicate.priority = 1;
  duplicate.id = "duplicate";
  harness.sheet.mutable_conditional_formats().push_back(MakeBlock({MakeRange(0, 0, 4, 0)}, {duplicate}));

  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 4, 0), MakeHost(harness));
  ASSERT_TRUE(static_cast<bool>(result));
  ASSERT_EQ(result.value().size(), 2u);
  EXPECT_EQ(result.value()[0].cell, At(0, 0));
  EXPECT_EQ(result.value()[1].cell, At(1, 0));
}

TEST(CFEvaluator, EvaluateCfForRangeFormulaCacheShiftsRelativeReferencePerCell) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(5.0));
  harness.sheet.set_cell_value(1, 0, Value::number(20.0));
  harness.sheet.set_cell_value(2, 0, Value::number(5.0));

  CFRule expression = MakeRule(RuleType::Expression);
  expression.priority = 1;
  expression.id = "relative";
  expression.formula1 = "A1>10";
  harness.sheet.mutable_conditional_formats().push_back(MakeBlock({MakeRange(0, 0, 2, 0)}, {std::move(expression)}));

  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 2, 0), MakeHost(harness));
  ASSERT_TRUE(static_cast<bool>(result));
  ASSERT_EQ(result.value().size(), 1u);
  EXPECT_EQ(result.value()[0].cell, At(1, 0));
  ASSERT_EQ(result.value()[0].matches.size(), 1u);
  EXPECT_EQ(result.value()[0].matches[0].rule_id, "relative");
}

TEST(CFEvaluator, EvaluateCfForRangeMalformedFormulaMatchesNoTarget) {
  CFEvalHarness harness;
  CFRule expression = MakeRule(RuleType::Expression);
  expression.priority = 1;
  expression.id = "malformed";
  expression.formula1 = "(((";
  harness.sheet.mutable_conditional_formats().push_back(MakeBlock({MakeRange(0, 0, 2, 0)}, {std::move(expression)}));

  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 2, 0), MakeHost(harness));
  ASSERT_TRUE(static_cast<bool>(result));
  EXPECT_TRUE(result.value().empty());
}

TEST(CFEvaluator, EvaluateCfForRangeFormulaCacheEvaluatesVolatileFunctionPerTarget) {
  CFEvalHarness harness;
  eval::FunctionRegistry registry;
  ASSERT_TRUE(registry.register_function(eval::FunctionDef{"CF_FORMULA_CALL", 0U, 0U, &CfFormulaCallCounter}));

  CFRule expression = MakeRule(RuleType::Expression);
  expression.priority = 1;
  expression.id = "per-target";
  expression.formula1 = "CF_FORMULA_CALL()";
  harness.sheet.mutable_conditional_formats().push_back(MakeBlock({MakeRange(0, 0, 2, 0)}, {std::move(expression)}));

  CFHost host = MakeHost(harness);
  host.registry = &registry;
  g_cf_formula_calls = 0;
  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 2, 0), host);
  ASSERT_TRUE(static_cast<bool>(result));
  ASSERT_EQ(result.value().size(), 3u);
  EXPECT_EQ(g_cf_formula_calls, 3U);
}

TEST(CFEvaluator, EvaluateCfForRangeFormulaCacheShiftsSameSourceForDistinctAnchors) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(1.0));  // A1 matches block A.
  harness.sheet.set_cell_value(1, 0, Value::number(0.0));  // A2 does not.
  harness.sheet.set_cell_value(0, 1, Value::number(0.0));  // B1's own value is not consulted by block B.
  harness.sheet.set_cell_value(1, 1, Value::number(1.0));  // B2 does not match A2 through block B.

  CFRule first = MakeRule(RuleType::Expression);
  first.priority = 1;
  first.id = "block-a";
  first.formula1 = "A1=1";
  CFRule second = MakeRule(RuleType::Expression);
  second.priority = 1;
  second.id = "block-b";
  second.formula1 = "A1=1";
  harness.sheet.mutable_conditional_formats().push_back(MakeBlock({MakeRange(0, 0, 1, 0)}, {std::move(first)}));
  harness.sheet.mutable_conditional_formats().push_back(MakeBlock({MakeRange(0, 1, 1, 1)}, {std::move(second)}));

  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 1, 1), MakeHost(harness));
  ASSERT_TRUE(static_cast<bool>(result));
  ASSERT_EQ(result.value().size(), 2u);
  EXPECT_EQ(result.value()[0].cell, At(0, 0));
  ASSERT_EQ(result.value()[0].matches.size(), 1u);
  EXPECT_EQ(result.value()[0].matches[0].rule_id, "block-a");
  EXPECT_EQ(result.value()[1].cell, At(0, 1));
  ASSERT_EQ(result.value()[1].matches.size(), 1u);
  EXPECT_EQ(result.value()[1].matches[0].rule_id, "block-b");
}

TEST(CFEvaluator, EvaluateCfForRangeFormulaCacheCoversCellIsFormulaPair) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(1.0));
  harness.sheet.set_cell_value(1, 0, Value::number(2.0));
  harness.sheet.set_cell_value(0, 1, Value::number(3.0));
  harness.sheet.set_cell_value(1, 1, Value::number(4.0));

  CFRule between = MakeRule(RuleType::CellIs);
  between.priority = 1;
  between.id = "between";
  between.op = CellIsOperator::Between;
  between.formula1 = "A1";
  between.formula2 = "B1";
  harness.sheet.mutable_conditional_formats().push_back(MakeBlock({MakeRange(0, 0, 1, 0)}, {std::move(between)}));

  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 1, 0), MakeHost(harness));
  ASSERT_TRUE(static_cast<bool>(result));
  ASSERT_EQ(result.value().size(), 2u);
  EXPECT_EQ(result.value()[0].cell, At(0, 0));
  EXPECT_EQ(result.value()[1].cell, At(1, 0));
}

TEST(CFEvaluator, EvaluateCfForRangeFormulaCacheCoversCfvoFormulaThresholds) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(10.0));
  harness.sheet.set_cell_value(1, 0, Value::number(20.0));

  CFRule color_scale = MakeRule(RuleType::ColorScale);
  color_scale.priority = 1;
  color_scale.id = "formula-scale";
  color_scale.color_scale =
      ColorScaleSpec{{CfValueObject{CfvoType::Formula, "A1", true}, CfValueObject{CfvoType::Formula, "A2", true}},
                     {RGB(255, 0, 0), RGB(0, 0, 255)}};
  harness.sheet.mutable_conditional_formats().push_back(MakeBlock({MakeRange(0, 0, 1, 0)}, {std::move(color_scale)}));

  const auto result = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 1, 0), MakeHost(harness));
  ASSERT_TRUE(static_cast<bool>(result));
  ASSERT_EQ(result.value().size(), 2u);
  ASSERT_TRUE(result.value()[0].matches[0].resolved_fill_color.has_value());
  ASSERT_TRUE(result.value()[1].matches[0].resolved_fill_color.has_value());
  EXPECT_EQ(result.value()[0].matches[0].resolved_fill_color->r, 255U);
  EXPECT_EQ(result.value()[1].matches[0].resolved_fill_color->b, 255U);
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
