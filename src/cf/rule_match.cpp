//
// Per-rule-kind matching for the CF evaluator. See rule_match.h /
// cf/cf_evaluator.h for the contract; the public `match_rule(...)`
// overloads are defined here.

#include "cf/rule_match.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string_view>

#include "cf/cf_evaluator.h"
#include "cf/cf_helpers.h"
#include "cf/cf_types.h"
#include "cf/scale_evaluator.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "value.h"

namespace formulon::cf {

namespace {

// ---------------------------------------------------------------------------
// CellIs — literal-only fast path (value-only overload).
// ---------------------------------------------------------------------------

bool match_cell_is(const CFRule& rule, const Value& cell_value) {
  if (!rule.op.has_value() || !rule.formula1.has_value()) {
    return false;
  }

  const auto operand1 = helpers::parse_literal(*rule.formula1);
  if (!operand1.has_value()) {
    return false;
  }

  const CellIsOperator cell_op = *rule.op;
  if (cell_op == CellIsOperator::Between || cell_op == CellIsOperator::NotBetween) {
    if (!rule.formula2.has_value()) {
      return false;
    }
    const auto operand2 = helpers::parse_literal(*rule.formula2);
    if (!operand2.has_value()) {
      return false;
    }
    const auto lo_cmp = helpers::compare_cell_to_literal(cell_value, *operand1);
    const auto hi_cmp = helpers::compare_cell_to_literal(cell_value, *operand2);
    if (!lo_cmp.has_value() || !hi_cmp.has_value()) {
      return false;
    }
    const bool inside = (*lo_cmp >= 0) && (*hi_cmp <= 0);
    return cell_op == CellIsOperator::Between ? inside : !inside;
  }

  const auto cmp = helpers::compare_cell_to_literal(cell_value, *operand1);
  if (!cmp.has_value()) {
    return false;
  }
  switch (cell_op) {
    case CellIsOperator::LessThan:
      return *cmp < 0;
    case CellIsOperator::LessThanOrEqual:
      return *cmp <= 0;
    case CellIsOperator::Equal:
      return *cmp == 0;
    case CellIsOperator::NotEqual:
      return *cmp != 0;
    case CellIsOperator::GreaterThanOrEqual:
      return *cmp >= 0;
    case CellIsOperator::GreaterThan:
      return *cmp > 0;
    case CellIsOperator::Between:
    case CellIsOperator::NotBetween:
      return false;  // Already handled above; here only for switch coverage.
  }
  return false;
}

// ---------------------------------------------------------------------------
// Text family — containsText / notContainsText / beginsWith / endsWith.
// ---------------------------------------------------------------------------

// Excel's text-family rules are built on the displayed text of the cell,
// not its raw kind. A number or blank cell is coerced to its General
// rendering (`512`, ``) before the substring test, matching Excel's
// generated SEARCH-based formula.
//
// The negative form (`NotContainsText`) is the predicate complement: a
// cell that does not contain the needle matches. A non-text cell never
// contains a non-empty needle, so the negation fires on numeric / blank
// cells — Excel highlights them. Error cells have no displayed text the
// rule can search, so the whole family is inert on them.
bool match_text_rule(const CFRule& rule, const Value& cell_value) {
  if (!rule.text.has_value()) {
    return false;
  }
  // Errors carry no searchable display text; neither the positive nor the
  // negative forms apply to them (Excel leaves error cells unhighlighted).
  if (cell_value.is_error()) {
    return false;
  }
  const auto coerced = eval::coerce_to_text(cell_value);
  if (!coerced.has_value()) {
    // Arrays / refs / lambdas have no scalar rendering; treat as inert.
    return false;
  }
  const std::string_view cell_text = coerced.value();
  const std::string_view needle = *rule.text;
  switch (rule.type) {
    case RuleType::ContainsText:
      return helpers::icase_contains(cell_text, needle);
    case RuleType::NotContainsText:
      return !helpers::icase_contains(cell_text, needle);
    case RuleType::BeginsWith:
      return helpers::icase_starts_with(cell_text, needle);
    case RuleType::EndsWith:
      return helpers::icase_ends_with(cell_text, needle);
    default:
      return false;
  }
}

// ---------------------------------------------------------------------------
// CellIs — formula-evaluator fallback (context-aware overload).
// ---------------------------------------------------------------------------

// Same shape as `match_cell_is`, but resolves non-literal formula1 /
// formula2 through the formula evaluator so cellIs rules with
// references or arithmetic operands work end-to-end.
bool match_cell_is_via_evaluator(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  if (!rule.op.has_value() || !rule.formula1.has_value()) {
    return false;
  }

  const auto operand1 = helpers::cell_is_operand(*rule.formula1, ctx);
  if (!operand1.has_value()) {
    return false;
  }

  const CellIsOperator cell_op = *rule.op;
  if (cell_op == CellIsOperator::Between || cell_op == CellIsOperator::NotBetween) {
    if (!rule.formula2.has_value()) {
      return false;
    }
    const auto operand2 = helpers::cell_is_operand(*rule.formula2, ctx);
    if (!operand2.has_value()) {
      return false;
    }
    const auto lo_cmp = helpers::compare_cell_to_literal(cell_value, *operand1);
    const auto hi_cmp = helpers::compare_cell_to_literal(cell_value, *operand2);
    if (!lo_cmp.has_value() || !hi_cmp.has_value()) {
      return false;
    }
    const bool inside = (*lo_cmp >= 0) && (*hi_cmp <= 0);
    return cell_op == CellIsOperator::Between ? inside : !inside;
  }

  const auto cmp = helpers::compare_cell_to_literal(cell_value, *operand1);
  if (!cmp.has_value()) {
    return false;
  }
  switch (cell_op) {
    case CellIsOperator::LessThan:
      return *cmp < 0;
    case CellIsOperator::LessThanOrEqual:
      return *cmp <= 0;
    case CellIsOperator::Equal:
      return *cmp == 0;
    case CellIsOperator::NotEqual:
      return *cmp != 0;
    case CellIsOperator::GreaterThanOrEqual:
      return *cmp >= 0;
    case CellIsOperator::GreaterThan:
      return *cmp > 0;
    case CellIsOperator::Between:
    case CellIsOperator::NotBetween:
      return false;  // Already handled above; here only for switch coverage.
  }
  return false;
}

// ---------------------------------------------------------------------------
// Expression rules.
// ---------------------------------------------------------------------------

// Excel's expression-rule truthiness: only a non-zero number or
// `TRUE` triggers a match. Text, errors, blanks, and arrays do not.
bool value_is_truthy(const Value& value) {
  if (value.is_boolean()) {
    return value.as_boolean();
  }
  if (value.is_number()) {
    return value.as_number() != 0.0;
  }
  return false;
}

bool match_expression(const CFRule& rule, const CFEvalContext& ctx) {
  if (!rule.formula1.has_value()) {
    return false;
  }
  const Value evaluated = helpers::parse_shift_evaluate(*rule.formula1, ctx);
  return value_is_truthy(evaluated);
}

// ---------------------------------------------------------------------------
// TimePeriod — bucket a date serial against `today_serial`.
//
// Excel's TimePeriod buckets are inclusive day-aligned ranges. The
// helpers floor to the integer serial (drop time-of-day) before any
// bucket comparison so a cell carrying `2024-03-15 13:30` matches the
// bucket for `2024-03-15` regardless of fractional time.
// ---------------------------------------------------------------------------

constexpr int kDaysPerWeek = 7;
constexpr double kLast7DaysSpan = 6.0;       // today - 6 ... today inclusive
constexpr double kWeekLastDayOffset = 6.0;   // Sunday + 6 = Saturday
constexpr double kPriorWeekDayOffset = 1.0;  // Sunday - 1 = prior Saturday

bool match_time_period(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  if (!rule.time_period.has_value() || !ctx.today_serial.has_value()) {
    return false;
  }
  if (!cell_value.is_number()) {
    return false;
  }
  const double cell_serial = std::floor(cell_value.as_number());
  if (cell_serial < 0.0) {
    return false;
  }
  const double today = std::floor(*ctx.today_serial);

  switch (*rule.time_period) {
    case TimePeriod::Today:
      return cell_serial == today;
    case TimePeriod::Yesterday:
      return cell_serial == today - 1.0;
    case TimePeriod::Tomorrow:
      return cell_serial == today + 1.0;
    case TimePeriod::Last7Days:
      return cell_serial >= today - kLast7DaysSpan && cell_serial <= today;
    case TimePeriod::ThisWeek: {
      const double sunday = helpers::sunday_of_week(today);
      return cell_serial >= sunday && cell_serial <= sunday + kWeekLastDayOffset;
    }
    case TimePeriod::LastWeek: {
      const double sunday = helpers::sunday_of_week(today);
      return cell_serial >= sunday - kDaysPerWeek && cell_serial <= sunday - kPriorWeekDayOffset;
    }
    case TimePeriod::NextWeek: {
      const double sunday = helpers::sunday_of_week(today);
      return cell_serial >= sunday + kDaysPerWeek && cell_serial <= sunday + kDaysPerWeek + kWeekLastDayOffset;
    }
    case TimePeriod::ThisMonth:
      return helpers::year_month_from_serial(cell_serial) == helpers::year_month_from_serial(today);
    case TimePeriod::LastMonth:
      return helpers::year_month_from_serial(cell_serial) ==
             helpers::shift_year_month(helpers::year_month_from_serial(today), -1);
    case TimePeriod::NextMonth:
      return helpers::year_month_from_serial(cell_serial) ==
             helpers::shift_year_month(helpers::year_month_from_serial(today), 1);
  }
  return false;
}

// ---------------------------------------------------------------------------
// DuplicateValues / UniqueValues — population statistics over the sqref.
// ---------------------------------------------------------------------------

bool match_duplicate_or_unique(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  if (ctx.sqref == nullptr || ctx.eval_ctx == nullptr || ctx.eval_ctx->current_sheet() == nullptr) {
    return false;
  }
  if (!cell_value.is_number() && !cell_value.is_boolean() && !cell_value.is_text()) {
    return false;
  }
  const std::size_t count = helpers::count_matches_in_sqref(cell_value, *ctx.sqref, *ctx.eval_ctx->current_sheet());
  switch (rule.type) {
    case RuleType::DuplicateValues:
      return count >= 2;
    case RuleType::UniqueValues:
      return count == 1;
    default:
      return false;
  }
}

// ---------------------------------------------------------------------------
// AboveAverage / Top10 — population statistics over the sqref.
//
// Both rules collect the numeric population from `ctx.sqref` (booleans
// and text are excluded — Excel folds them into the rule via SUMPRODUCT
// in the generated formula, but the closure harness will refine the
// stance when oracle data lands). Errors and blanks never participate.
// ---------------------------------------------------------------------------

constexpr double kPercentDivisor = 100.0;
constexpr std::int32_t kDefaultTop10Rank = 10;

bool match_above_average(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  ColorScalePopulation fallback;
  const ColorScalePopulation* pop = helpers::numeric_population_for_cell(cell_value, ctx, fallback);
  if (pop == nullptr) {
    return false;
  }
  const double mean = helpers::mean_of(pop->sorted);

  double threshold = mean;
  if (rule.std_dev.has_value() && *rule.std_dev > 0.0) {
    const double sigma = helpers::sample_stddev(pop->sorted, mean);
    const double offset = *rule.std_dev * sigma;
    threshold = rule.above_average ? mean + offset : mean - offset;
  }

  const double cell = cell_value.as_number();
  if (rule.above_average) {
    return rule.equal_average ? cell >= threshold : cell > threshold;
  }
  return rule.equal_average ? cell <= threshold : cell < threshold;
}

// Resolves the rank index `n` for Top10 — the cell-count to highlight
// before tie-inclusion expands the matched set. Excel clamps to
// `[1, count]` and uses floor-truncation when interpreting `rank` as a
// percent of the population.
std::size_t resolve_top10_rank(const CFRule& rule, std::size_t population) {
  const std::int32_t raw_rank = rule.rank.value_or(kDefaultTop10Rank);
  if (raw_rank <= 0 || population == 0) {
    return 0;
  }
  std::int64_t resolved = raw_rank;
  if (rule.percent) {
    resolved = static_cast<std::int64_t>(
        std::floor(static_cast<double>(population) * static_cast<double>(raw_rank) / kPercentDivisor));
  }
  if (resolved < 1) {
    resolved = 1;
  }
  if (resolved > static_cast<std::int64_t>(population)) {
    resolved = static_cast<std::int64_t>(population);
  }
  return static_cast<std::size_t>(resolved);
}

bool match_top10(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  ColorScalePopulation fallback;
  const ColorScalePopulation* pop = helpers::numeric_population_for_cell(cell_value, ctx, fallback);
  if (pop == nullptr) {
    return false;
  }
  const std::size_t rank_n = resolve_top10_rank(rule, pop->sorted.size());
  if (rank_n == 0) {
    return false;
  }

  // Sorted ascending: bottom-N threshold is `sorted[rank_n - 1]` (the
  // n-th smallest); top-N threshold is `sorted[count - rank_n]` (the
  // n-th largest). Tie inclusion comes from `<=` / `>=` against the
  // threshold so cells equal to the rank cutoff still match.
  const double cell = cell_value.as_number();
  if (rule.bottom) {
    const double threshold = pop->sorted[rank_n - 1];
    return cell <= threshold;
  }
  const double threshold = pop->sorted[pop->sorted.size() - rank_n];
  return cell >= threshold;
}

}  // namespace

// ---------------------------------------------------------------------------
// Public `match_rule` overloads.
// ---------------------------------------------------------------------------

bool match_rule(const CFRule& rule, const Value& cell_value) {
  switch (rule.type) {
    case RuleType::ContainsBlanks:
      return cell_value.is_blank();
    case RuleType::NotContainsBlanks:
      return !cell_value.is_blank();
    case RuleType::ContainsErrors:
      return cell_value.is_error();
    case RuleType::NotContainsErrors:
      return !cell_value.is_error();
    case RuleType::CellIs:
      return match_cell_is(rule, cell_value);
    case RuleType::ContainsText:
    case RuleType::NotContainsText:
    case RuleType::BeginsWith:
    case RuleType::EndsWith:
      return match_text_rule(rule, cell_value);
    // Rule types whose evaluator lands in subsequent PRs return false
    // here so a caller that walks the full rule list does not mis-fire
    // on a partially-implemented engine. The UI is expected to gate on
    // the engine version it links against.
    case RuleType::Expression:
    case RuleType::ColorScale:
    case RuleType::DataBar:
    case RuleType::IconSet:
    case RuleType::Top10:
    case RuleType::AboveAverage:
    case RuleType::TimePeriod:
    case RuleType::DuplicateValues:
    case RuleType::UniqueValues:
      return false;
  }
  return false;
}

bool match_rule(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  switch (rule.type) {
    case RuleType::Expression:
      return match_expression(rule, ctx);
    case RuleType::CellIs:
      return match_cell_is_via_evaluator(rule, cell_value, ctx);
    case RuleType::TimePeriod:
      return match_time_period(rule, cell_value, ctx);
    case RuleType::DuplicateValues:
    case RuleType::UniqueValues:
      return match_duplicate_or_unique(rule, cell_value, ctx);
    case RuleType::AboveAverage:
      return match_above_average(rule, cell_value, ctx);
    case RuleType::Top10:
      return match_top10(rule, cell_value, ctx);
    case RuleType::ColorScale:
      return scales::match_color_scale(rule, cell_value, ctx);
    case RuleType::DataBar:
      return scales::match_data_bar(rule, cell_value, ctx);
    case RuleType::IconSet:
      return scales::match_icon_set(rule, cell_value, ctx);
    case RuleType::ContainsBlanks:
    case RuleType::NotContainsBlanks:
    case RuleType::ContainsErrors:
    case RuleType::NotContainsErrors:
    case RuleType::ContainsText:
    case RuleType::NotContainsText:
    case RuleType::BeginsWith:
    case RuleType::EndsWith:
      // Value-only and not-yet-implemented rule types delegate to the
      // simple overload, which already encodes their semantics (or
      // false-fallthrough for the unimplemented kinds).
      return match_rule(rule, cell_value);
  }
  return false;
}

}  // namespace formulon::cf
