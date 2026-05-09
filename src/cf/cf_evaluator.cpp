// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the CF evaluator. See cf_evaluator.h for the
// scoped contract and the staged-PR roadmap.

#include "cf/cf_evaluator.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <functional>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "eval/date_time.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "parser/ast.h"
#include "parser/ast_shift.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon::cf {

// File-scope (non-anonymous) so the type matches the forward
// declaration in `cf_evaluator.h` and `CFEvalContext::cached_population`
// can carry a pointer to it. Layout is private to this translation
// unit; external callers see only the forward decl.
struct ColorScalePopulation {
  std::vector<double> sorted;  // ascending
  double min = 0.0;
  double max = 0.0;
};

namespace {

// Small literal-only operand used by the `cellIs` evaluator. Owns its
// text payload (unlike `Value::Text`, which is a non-owning view), so
// it can survive the parse helper's stack frame without an arena.
struct LiteralOperand {
  enum class Kind : std::uint8_t { Number, Bool, Text };
  Kind kind = Kind::Number;
  double number_value = 0.0;
  bool bool_value = false;
  std::string text_value;
};

bool icase_equal(std::string_view lhs, std::string_view rhs) {
  if (lhs.size() != rhs.size()) {
    return false;
  }
  for (std::size_t i = 0; i < lhs.size(); ++i) {
    char left_ch = lhs[i];
    char right_ch = rhs[i];
    if (left_ch >= 'A' && left_ch <= 'Z') {
      left_ch = static_cast<char>(left_ch - 'A' + 'a');
    }
    if (right_ch >= 'A' && right_ch <= 'Z') {
      right_ch = static_cast<char>(right_ch - 'A' + 'a');
    }
    if (left_ch != right_ch) {
      return false;
    }
  }
  return true;
}

// Three-way ASCII case-insensitive compare. Sufficient for cellIs text
// equality / ordering on the literals Excel emits; full Unicode-aware
// case-folding belongs with the broader text-comparison story.
int icase_compare(std::string_view lhs, std::string_view rhs) {
  const std::size_t shorter = std::min(lhs.size(), rhs.size());
  for (std::size_t i = 0; i < shorter; ++i) {
    char left_ch = lhs[i];
    char right_ch = rhs[i];
    if (left_ch >= 'A' && left_ch <= 'Z') {
      left_ch = static_cast<char>(left_ch - 'A' + 'a');
    }
    if (right_ch >= 'A' && right_ch <= 'Z') {
      right_ch = static_cast<char>(right_ch - 'A' + 'a');
    }
    if (left_ch != right_ch) {
      return left_ch < right_ch ? -1 : 1;
    }
  }
  if (lhs.size() == rhs.size()) {
    return 0;
  }
  return lhs.size() < rhs.size() ? -1 : 1;
}

// Parses a `cellIs` formula source as a literal. Excel emits cellIs
// `formula1` / `formula2` as either a bare number (`10`, `-3.5`), a
// boolean keyword (`TRUE` / `FALSE`), or a quoted string with `""`
// double-quote escapes. Anything else (a reference, an arithmetic
// expression) returns nullopt; PR8 brings in the formula evaluator
// so those operand shapes can be supported.
std::optional<LiteralOperand> parse_literal(const std::string& source) {
  if (source.empty()) {
    return std::nullopt;
  }

  // Quoted text literal with `""` -> `"` unescape.
  if (source.size() >= 2 && source.front() == '"' && source.back() == '"') {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Text;
    std::string& out = operand.text_value;
    out.reserve(source.size() - 2);
    for (std::size_t i = 1; i + 1 < source.size(); ++i) {
      if (source[i] == '"' && i + 2 < source.size() && source[i + 1] == '"') {
        out.push_back('"');
        ++i;
      } else {
        out.push_back(source[i]);
      }
    }
    return operand;
  }

  // Boolean keywords (case-insensitive, matching Excel's tolerance).
  if (icase_equal(source, "TRUE")) {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Bool;
    operand.bool_value = true;
    return operand;
  }
  if (icase_equal(source, "FALSE")) {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Bool;
    operand.bool_value = false;
    return operand;
  }

  // Numeric literal — consumed in full or rejected.
  char* end = nullptr;
  const double parsed = std::strtod(source.c_str(), &end);
  if (end != source.c_str() && *end == '\0') {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Number;
    operand.number_value = parsed;
    return operand;
  }
  return std::nullopt;
}

double literal_as_number(const LiteralOperand& operand) {
  switch (operand.kind) {
    case LiteralOperand::Kind::Number:
      return operand.number_value;
    case LiteralOperand::Kind::Bool:
      return operand.bool_value ? 1.0 : 0.0;
    case LiteralOperand::Kind::Text:
      return 0.0;  // Caller must check kinds before reaching this branch.
  }
  return 0.0;
}

// Three-way compare of `cell` against a literal operand. Returns
// nullopt when the kinds are incompatible (e.g. text rule against a
// numeric cell, any rule against an error cell). Mirroring Excel's
// behaviour on cross-kind cellIs comparisons in full needs oracle
// data; PR7 takes the conservative stance that incompatible kinds
// don't compare and lands the refinement when the closure harness
// covers cellIs.
std::optional<int> compare_cell_to_literal(const Value& cell, const LiteralOperand& operand) {
  if (cell.is_error() || cell.is_blank()) {
    return std::nullopt;
  }

  if (operand.kind == LiteralOperand::Kind::Text) {
    if (!cell.is_text()) {
      return std::nullopt;
    }
    return icase_compare(cell.as_text(), operand.text_value);
  }

  // Numeric / bool comparison: lift bool to 0/1.
  double cell_num = 0.0;
  if (cell.is_number()) {
    cell_num = cell.as_number();
  } else if (cell.is_boolean()) {
    cell_num = cell.as_boolean() ? 1.0 : 0.0;
  } else {
    return std::nullopt;
  }
  const double lit_num = literal_as_number(operand);
  if (cell_num < lit_num) {
    return -1;
  }
  if (cell_num > lit_num) {
    return 1;
  }
  return 0;
}

// Folds an evaluated `Value` into the literal-operand machinery so the
// existing comparison helpers can stay the only path. Errors, blanks,
// arrays, refs, and lambdas have no useful CellIs interpretation; they
// produce nullopt and the caller treats the rule as not-matching.
std::optional<LiteralOperand> value_to_operand(const Value& evaluated) {
  if (evaluated.is_number()) {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Number;
    operand.number_value = evaluated.as_number();
    return operand;
  }
  if (evaluated.is_boolean()) {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Bool;
    operand.bool_value = evaluated.as_boolean();
    return operand;
  }
  if (evaluated.is_text()) {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Text;
    operand.text_value = std::string(evaluated.as_text());
    return operand;
  }
  return std::nullopt;
}

// Parses, shifts, and evaluates `source` as the body of a CF formula
// anchored at `ctx.anchor` and applied at `ctx.target`. Parser errors,
// arena exhaustion, and out-of-bounds shifts surface as Excel error
// values so the caller can decide how to interpret them.
Value parse_shift_evaluate(const std::string& source, const CFEvalContext& ctx) {
  if (ctx.arena == nullptr || ctx.registry == nullptr || ctx.eval_ctx == nullptr) {
    return Value::error(ErrorCode::Name);
  }

  // OOXML stores cellIs / expression `formula1` without a leading `=`;
  // the parser is happy with that, so no stripping is required.
  parser::Parser parser(source, *ctx.arena);
  const parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    return Value::error(ErrorCode::Name);
  }

  const std::int32_t row_delta = static_cast<std::int32_t>(ctx.target.row) - static_cast<std::int32_t>(ctx.anchor.row);
  const std::int32_t col_delta = static_cast<std::int32_t>(ctx.target.col) - static_cast<std::int32_t>(ctx.anchor.col);

  const parser::AstNode* shifted = parser::shift_relative_refs(*root, *ctx.arena, row_delta, col_delta);
  if (shifted == nullptr) {
    return Value::error(ErrorCode::Name);
  }

  const eval::EvalContext target_ctx = ctx.eval_ctx->with_formula_cell(ctx.target.row, ctx.target.col);
  return eval::evaluate(*shifted, *ctx.arena, *ctx.registry, target_ctx);
}

// Returns the operand for a `cellIs` formula source, preferring the
// literal-parser fast path (no parser instantiation, no arena traffic)
// and falling back to the formula evaluator when the source isn't a
// bare literal.
std::optional<LiteralOperand> cell_is_operand(const std::string& source, const CFEvalContext& ctx) {
  auto literal = parse_literal(source);
  if (literal.has_value()) {
    return literal;
  }
  const Value evaluated = parse_shift_evaluate(source, ctx);
  return value_to_operand(evaluated);
}

bool match_cell_is(const CFRule& rule, const Value& cell_value) {
  if (!rule.op.has_value() || !rule.formula1.has_value()) {
    return false;
  }

  const auto operand1 = parse_literal(*rule.formula1);
  if (!operand1.has_value()) {
    return false;
  }

  const CellIsOperator cell_op = *rule.op;
  if (cell_op == CellIsOperator::Between || cell_op == CellIsOperator::NotBetween) {
    if (!rule.formula2.has_value()) {
      return false;
    }
    const auto operand2 = parse_literal(*rule.formula2);
    if (!operand2.has_value()) {
      return false;
    }
    const auto lo_cmp = compare_cell_to_literal(cell_value, *operand1);
    const auto hi_cmp = compare_cell_to_literal(cell_value, *operand2);
    if (!lo_cmp.has_value() || !hi_cmp.has_value()) {
      return false;
    }
    const bool inside = (*lo_cmp >= 0) && (*hi_cmp <= 0);
    return cell_op == CellIsOperator::Between ? inside : !inside;
  }

  const auto cmp = compare_cell_to_literal(cell_value, *operand1);
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

// ASCII case-insensitive substring containment. Returns `true` for an
// empty needle (every text contains the empty string), matching the
// SEARCH/`<containsText>` convention Excel uses to compile these rules.
bool icase_contains(std::string_view haystack, std::string_view needle) {
  if (needle.empty()) {
    return true;
  }
  if (needle.size() > haystack.size()) {
    return false;
  }
  const std::size_t span = haystack.size() - needle.size();
  for (std::size_t i = 0; i <= span; ++i) {
    bool matched = true;
    for (std::size_t j = 0; j < needle.size(); ++j) {
      if (strings::ascii_to_lower(haystack[i + j]) != strings::ascii_to_lower(needle[j])) {
        matched = false;
        break;
      }
    }
    if (matched) {
      return true;
    }
  }
  return false;
}

bool icase_starts_with(std::string_view text, std::string_view prefix) {
  if (prefix.size() > text.size()) {
    return false;
  }
  return strings::case_insensitive_eq(text.substr(0, prefix.size()), prefix);
}

bool icase_ends_with(std::string_view text, std::string_view suffix) {
  if (suffix.size() > text.size()) {
    return false;
  }
  return strings::case_insensitive_eq(text.substr(text.size() - suffix.size()), suffix);
}

// Evaluates the `containsText` / `notContainsText` / `beginsWith` /
// `endsWith` family. Conservative cross-kind stance: if the cell isn't
// text, no rule in the family matches (including the negative
// `NotContainsText`). Excel implicitly coerces non-text cells via the
// generated SEARCH-based formula, but folding that in needs oracle
// data; the closure harness will refine when it lands.
bool match_text_rule(const CFRule& rule, const Value& cell_value) {
  if (!rule.text.has_value() || !cell_value.is_text()) {
    return false;
  }
  const std::string_view cell_text = cell_value.as_text();
  const std::string_view needle = *rule.text;
  switch (rule.type) {
    case RuleType::ContainsText:
      return icase_contains(cell_text, needle);
    case RuleType::NotContainsText:
      return !icase_contains(cell_text, needle);
    case RuleType::BeginsWith:
      return icase_starts_with(cell_text, needle);
    case RuleType::EndsWith:
      return icase_ends_with(cell_text, needle);
    default:
      return false;
  }
}

}  // namespace

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

namespace {

// Same shape as `match_cell_is`, but resolves non-literal formula1 /
// formula2 through the formula evaluator so cellIs rules with
// references or arithmetic operands work end-to-end.
bool match_cell_is_via_evaluator(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  if (!rule.op.has_value() || !rule.formula1.has_value()) {
    return false;
  }

  const auto operand1 = cell_is_operand(*rule.formula1, ctx);
  if (!operand1.has_value()) {
    return false;
  }

  const CellIsOperator cell_op = *rule.op;
  if (cell_op == CellIsOperator::Between || cell_op == CellIsOperator::NotBetween) {
    if (!rule.formula2.has_value()) {
      return false;
    }
    const auto operand2 = cell_is_operand(*rule.formula2, ctx);
    if (!operand2.has_value()) {
      return false;
    }
    const auto lo_cmp = compare_cell_to_literal(cell_value, *operand1);
    const auto hi_cmp = compare_cell_to_literal(cell_value, *operand2);
    if (!lo_cmp.has_value() || !hi_cmp.has_value()) {
      return false;
    }
    const bool inside = (*lo_cmp >= 0) && (*hi_cmp <= 0);
    return cell_op == CellIsOperator::Between ? inside : !inside;
  }

  const auto cmp = compare_cell_to_literal(cell_value, *operand1);
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
  const Value evaluated = parse_shift_evaluate(*rule.formula1, ctx);
  return value_is_truthy(evaluated);
}

// ---------------------------------------------------------------------------
// TimePeriod — bucket a date serial against `today_serial`.
//
// Excel's TimePeriod buckets are inclusive day-aligned ranges. The
// helpers below floor to the integer serial (drop time-of-day) before
// any bucket comparison so a cell carrying `2024-03-15 13:30` matches
// the bucket for `2024-03-15` regardless of fractional time.
// ---------------------------------------------------------------------------

constexpr int kMonthsPerYear = 12;
constexpr int kDaysPerWeek = 7;
constexpr double kLast7DaysSpan = 6.0;       // today - 6 ... today inclusive
constexpr double kWeekLastDayOffset = 6.0;   // Sunday + 6 = Saturday
constexpr double kPriorWeekDayOffset = 1.0;  // Sunday - 1 = prior Saturday

// Excel weekday with `WEEKDAY(date, 1)` semantics: Sunday = 1, Saturday
// = 7. Computed from the proleptic-Gregorian day count produced by
// `days_from_civil`, which is bug-free; the 1900 leap-year ghost day
// is absorbed by `ymd_from_serial` upstream.
int weekday_sunday_one(double serial_floor) {
  const eval::date_time::YMD ymd = eval::date_time::ymd_from_serial(serial_floor);
  const std::int64_t days = eval::date_time::days_from_civil(ymd.y, ymd.m, ymd.d);
  // 1970-01-01 was a Thursday → Excel weekday 5. Adjust so days = 0
  // maps to 5, then take mod 7 and shift to the 1..7 range.
  const std::int64_t adjusted = ((days % kDaysPerWeek) + kDaysPerWeek + 4) % kDaysPerWeek;
  return static_cast<int>(adjusted) + 1;
}

// Serial of the Sunday opening the week that contains `serial_floor`.
double sunday_of_week(double serial_floor) {
  const int weekday = weekday_sunday_one(serial_floor);
  return serial_floor - (weekday - 1);
}

struct YearMonth {
  int year;
  unsigned month;  // 1..12

  friend bool operator==(YearMonth lhs, YearMonth rhs) noexcept {
    return lhs.year == rhs.year && lhs.month == rhs.month;
  }
};

YearMonth year_month_from_serial(double serial_floor) {
  const eval::date_time::YMD ymd = eval::date_time::ymd_from_serial(serial_floor);
  return {ymd.y, ymd.m};
}

YearMonth shift_year_month(YearMonth anchor, int delta_months) {
  const int month_index = static_cast<int>(anchor.month) - 1 + delta_months;
  // Floor-divide month_index by 12 so negative offsets borrow a year
  // correctly (e.g. month_index = -1 → year - 1, month = 12).
  int year_delta = month_index / kMonthsPerYear;
  int normalised_index = month_index % kMonthsPerYear;
  if (normalised_index < 0) {
    normalised_index += kMonthsPerYear;
    --year_delta;
  }
  return {anchor.year + year_delta, static_cast<unsigned>(normalised_index + 1)};
}

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
      const double sunday = sunday_of_week(today);
      return cell_serial >= sunday && cell_serial <= sunday + kWeekLastDayOffset;
    }
    case TimePeriod::LastWeek: {
      const double sunday = sunday_of_week(today);
      return cell_serial >= sunday - kDaysPerWeek && cell_serial <= sunday - kPriorWeekDayOffset;
    }
    case TimePeriod::NextWeek: {
      const double sunday = sunday_of_week(today);
      return cell_serial >= sunday + kDaysPerWeek && cell_serial <= sunday + kDaysPerWeek + kWeekLastDayOffset;
    }
    case TimePeriod::ThisMonth:
      return year_month_from_serial(cell_serial) == year_month_from_serial(today);
    case TimePeriod::LastMonth:
      return year_month_from_serial(cell_serial) == shift_year_month(year_month_from_serial(today), -1);
    case TimePeriod::NextMonth:
      return year_month_from_serial(cell_serial) == shift_year_month(year_month_from_serial(today), 1);
  }
  return false;
}

// ---------------------------------------------------------------------------
// DuplicateValues / UniqueValues — population statistics over the sqref.
// ---------------------------------------------------------------------------

// Cross-kind=false equality used for duplicate detection. Numbers
// compare by exact IEEE-754 equality (Excel does not deduplicate by an
// epsilon — values that round to the same display string but differ in
// the binary representation count as distinct). Text uses ASCII
// case-insensitive equality, mirroring the cellIs / containsText
// stance. Booleans compare by identity. Errors and blanks never
// participate in deduplication: any pairing involving them is
// considered unequal so a sqref full of `#N/A` produces no
// `DuplicateValues` matches.
bool cf_values_equal(const Value& lhs, const Value& rhs) {
  if (lhs.is_number() && rhs.is_number()) {
    return lhs.as_number() == rhs.as_number();
  }
  if (lhs.is_boolean() && rhs.is_boolean()) {
    return lhs.as_boolean() == rhs.as_boolean();
  }
  if (lhs.is_text() && rhs.is_text()) {
    return strings::case_insensitive_eq(lhs.as_text(), rhs.as_text());
  }
  return false;
}

// Counts how many cells inside `sqref` (read spill-aware via
// `Sheet::resolve_cell_value`) carry a value equal to `target` under
// `cf_values_equal`. Cells whose value is not number / boolean / text
// are skipped at the equality layer — they cannot match `target` (and
// `target` itself is rejected upstream when its kind is unsupported).
std::size_t count_matches_in_sqref(const Value& target, const std::vector<CFCellRange>& sqref, const Sheet& sheet) {
  std::size_t count = 0;
  for (const CFCellRange& range : sqref) {
    for (std::uint32_t row = range.first.row; row <= range.last.row; ++row) {
      for (std::uint32_t col = range.first.col; col <= range.last.col; ++col) {
        const Value cell = sheet.resolve_cell_value(row, col);
        if (cf_values_equal(target, cell)) {
          ++count;
        }
      }
    }
  }
  return count;
}

bool match_duplicate_or_unique(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  if (ctx.sqref == nullptr || ctx.eval_ctx == nullptr || ctx.eval_ctx->current_sheet() == nullptr) {
    return false;
  }
  if (!cell_value.is_number() && !cell_value.is_boolean() && !cell_value.is_text()) {
    return false;
  }
  const std::size_t count = count_matches_in_sqref(cell_value, *ctx.sqref, *ctx.eval_ctx->current_sheet());
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

// Forward declarations: `ensure_population` and `gather_population` are
// defined alongside `ColorScalePopulation` in the ColorScale section
// below, but the AboveAverage / Top10 helpers above them want to share
// the same cached-or-compute path.
ColorScalePopulation gather_population(const std::vector<CFCellRange>& sqref, const Sheet& sheet);
const ColorScalePopulation& ensure_population(const CFEvalContext& ctx, ColorScalePopulation& fallback);
const ColorScalePopulation* numeric_population_for_cell(const Value& cell_value, const CFEvalContext& ctx,
                                                        ColorScalePopulation& fallback);

// Collects every numeric value in `sqref` (read spill-aware via
// `Sheet::resolve_cell_value`). Booleans and text are skipped — Excel's
// AboveAverage / Top10 use the numeric-only population; a closure-driven
// refinement may widen this if oracle data shows otherwise.
std::vector<double> collect_numeric_values(const std::vector<CFCellRange>& sqref, const Sheet& sheet) {
  std::vector<double> values;
  for (const CFCellRange& range : sqref) {
    for (std::uint32_t row = range.first.row; row <= range.last.row; ++row) {
      for (std::uint32_t col = range.first.col; col <= range.last.col; ++col) {
        const Value cell = sheet.resolve_cell_value(row, col);
        if (cell.is_number()) {
          values.push_back(cell.as_number());
        }
      }
    }
  }
  return values;
}

double mean_of(const std::vector<double>& values) {
  double sum = 0.0;
  for (double sample : values) {
    sum += sample;
  }
  return sum / static_cast<double>(values.size());
}

// Sample standard deviation (Bessel-corrected, n-1). Returns 0 when
// fewer than 2 numeric values are present so the threshold collapses
// to the mean and the rule still produces a well-defined match.
double sample_stddev(const std::vector<double>& values, double mean) {
  if (values.size() < 2) {
    return 0.0;
  }
  double sumsq = 0.0;
  for (double sample : values) {
    const double delta = sample - mean;
    sumsq += delta * delta;
  }
  return std::sqrt(sumsq / static_cast<double>(values.size() - 1));
}

bool match_above_average(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  ColorScalePopulation fallback;
  const ColorScalePopulation* pop = numeric_population_for_cell(cell_value, ctx, fallback);
  if (pop == nullptr) {
    return false;
  }
  const double mean = mean_of(pop->sorted);

  double threshold = mean;
  if (rule.std_dev.has_value() && *rule.std_dev > 0.0) {
    const double sigma = sample_stddev(pop->sorted, mean);
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
  const ColorScalePopulation* pop = numeric_population_for_cell(cell_value, ctx, fallback);
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

// ---------------------------------------------------------------------------
// ColorScale — resolve `<cfvo>` thresholds against the sqref population
// and linearly interpolate the bounding stop colours in RGB space.
// ---------------------------------------------------------------------------

constexpr double kColorChannelMax = 255.0;

ColorScalePopulation gather_population(const std::vector<CFCellRange>& sqref, const Sheet& sheet) {
  std::vector<double> values = collect_numeric_values(sqref, sheet);
  std::sort(values.begin(), values.end());
  ColorScalePopulation pop;
  if (!values.empty()) {
    pop.min = values.front();
    pop.max = values.back();
  }
  pop.sorted = std::move(values);
  return pop;
}

// Returns the cached population from `ctx` when present; otherwise
// gathers from the sheet into `fallback` and returns it. The returned
// reference is valid for as long as `fallback` (or the cache pointer)
// outlives the call site. Callers stack-allocate `fallback`.
const ColorScalePopulation& ensure_population(const CFEvalContext& ctx, ColorScalePopulation& fallback) {
  if (ctx.cached_population != nullptr) {
    return *ctx.cached_population;
  }
  fallback = gather_population(*ctx.sqref, *ctx.eval_ctx->current_sheet());
  return fallback;
}

const ColorScalePopulation* numeric_population_for_cell(const Value& cell_value, const CFEvalContext& ctx,
                                                        ColorScalePopulation& fallback) {
  if (ctx.sqref == nullptr || ctx.eval_ctx == nullptr || ctx.eval_ctx->current_sheet() == nullptr) {
    return nullptr;
  }
  if (!cell_value.is_number()) {
    return nullptr;
  }
  const ColorScalePopulation& pop = ensure_population(ctx, fallback);
  if (pop.sorted.empty()) {
    return nullptr;
  }
  return &pop;
}

// PERCENTILE.INC: linear interpolation between sorted points.
// `percentile` is 0..1. For empty populations the caller short-circuits.
double percentile_inc(const std::vector<double>& sorted, double percentile) {
  const std::size_t count = sorted.size();
  if (count == 1) {
    return sorted.front();
  }
  const double position = percentile * static_cast<double>(count - 1);
  const auto lower_index = static_cast<std::size_t>(std::floor(position));
  if (lower_index + 1 >= count) {
    return sorted.back();
  }
  const double fraction = position - static_cast<double>(lower_index);
  return sorted[lower_index] + fraction * (sorted[lower_index + 1] - sorted[lower_index]);
}

std::optional<double> parse_double(std::string_view source) {
  if (source.empty()) {
    return std::nullopt;
  }
  const std::string copy(source);
  char* end = nullptr;
  const double parsed = std::strtod(copy.c_str(), &end);
  if (end != copy.c_str() && *end == '\0') {
    return parsed;
  }
  return std::nullopt;
}

// Resolves a single `<cfvo>` to its threshold value. `Formula` CFVOs
// run through the same parse-shift-evaluate path the rest of the
// context-aware evaluator uses, anchored at the rule's anchor (the
// formula authoring cell). Returns nullopt when the CFVO cannot be
// resolved (e.g. malformed literal, formula evaluation error).
std::optional<double> resolve_cfvo(const CfValueObject& cfvo, const ColorScalePopulation& pop,
                                   const CFEvalContext& ctx) {
  switch (cfvo.type) {
    case CfvoType::Number:
      return parse_double(cfvo.value);
    case CfvoType::Percent: {
      auto pct = parse_double(cfvo.value);
      if (!pct.has_value()) {
        return std::nullopt;
      }
      return pop.min + (*pct / kPercentDivisor) * (pop.max - pop.min);
    }
    case CfvoType::Percentile: {
      auto pct = parse_double(cfvo.value);
      if (!pct.has_value() || pop.sorted.empty()) {
        return std::nullopt;
      }
      return percentile_inc(pop.sorted, *pct / kPercentDivisor);
    }
    case CfvoType::Min:
    case CfvoType::AutoMin:
      return pop.sorted.empty() ? std::optional<double>() : std::optional<double>(pop.min);
    case CfvoType::Max:
    case CfvoType::AutoMax:
      return pop.sorted.empty() ? std::optional<double>() : std::optional<double>(pop.max);
    case CfvoType::Formula: {
      const Value evaluated = parse_shift_evaluate(cfvo.value, ctx);
      if (evaluated.is_number()) {
        return evaluated.as_number();
      }
      if (evaluated.is_boolean()) {
        return evaluated.as_boolean() ? 1.0 : 0.0;
      }
      return std::nullopt;
    }
  }
  return std::nullopt;
}

std::optional<std::vector<double>> resolve_cfvo_list(const std::vector<CfValueObject>& cfvos,
                                                     const ColorScalePopulation& pop, const CFEvalContext& ctx) {
  std::vector<double> resolved;
  resolved.reserve(cfvos.size());
  for (const CfValueObject& cfvo : cfvos) {
    auto value = resolve_cfvo(cfvo, pop, ctx);
    if (!value.has_value()) {
      return std::nullopt;
    }
    resolved.push_back(*value);
  }
  return resolved;
}

// Linear interpolation between two sRGB colours. `fraction` is clamped
// to [0, 1] by the caller. Alpha is interpolated alongside RGB so
// stops with transparent components blend correctly.
Color interpolate_color(Color start, Color end, double fraction) {
  const auto blend = [fraction](std::uint8_t low_channel, std::uint8_t high_channel) {
    const double mixed = static_cast<double>(low_channel) +
                         fraction * (static_cast<double>(high_channel) - static_cast<double>(low_channel));
    const double clamped = std::max(0.0, std::min(kColorChannelMax, mixed));
    return static_cast<std::uint8_t>(std::lround(clamped));
  };
  Color out;
  out.r = blend(start.r, end.r);
  out.g = blend(start.g, end.g);
  out.b = blend(start.b, end.b);
  out.a = blend(start.a, end.a);
  return out;
}

// Returns the resolved cell colour for a `colorScale` rule applied to
// a numeric `cell_value`, honouring 2-stop and 3-stop scales. Returns
// nullopt for empty populations, malformed thresholds, or non-numeric
// cells (the rule still "applies", but the cell renders without fill).
std::optional<Color> resolve_color_scale(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  if (!rule.color_scale.has_value()) {
    return std::nullopt;
  }
  const ColorScaleSpec& spec = *rule.color_scale;
  if (spec.thresholds.size() != spec.colors.size() || spec.thresholds.size() < 2) {
    return std::nullopt;
  }

  ColorScalePopulation fallback;
  const ColorScalePopulation* pop = numeric_population_for_cell(cell_value, ctx, fallback);
  if (pop == nullptr) {
    return std::nullopt;
  }

  auto resolved_thresholds = resolve_cfvo_list(spec.thresholds, *pop, ctx);
  if (!resolved_thresholds.has_value()) {
    return std::nullopt;
  }

  const double cell = cell_value.as_number();
  // Locate the segment that contains the cell value. Cells outside the
  // outermost stops clamp to the boundary colour.
  if (cell <= resolved_thresholds->front()) {
    return spec.colors.front();
  }
  if (cell >= resolved_thresholds->back()) {
    return spec.colors.back();
  }
  for (std::size_t i = 0; i + 1 < resolved_thresholds->size(); ++i) {
    const double lower_bound = (*resolved_thresholds)[i];
    const double upper_bound = (*resolved_thresholds)[i + 1];
    if (cell >= lower_bound && cell <= upper_bound) {
      // When the segment collapses (lower == upper), pick the upper
      // colour; the cell is exactly at a stop so either end is correct.
      const double span = upper_bound - lower_bound;
      const double fraction = span == 0.0 ? 1.0 : (cell - lower_bound) / span;
      return interpolate_color(spec.colors[i], spec.colors[i + 1], fraction);
    }
  }
  return spec.colors.back();
}

bool match_color_scale(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  // ColorScale "matches" any cell for which a fill colour can be
  // computed. Non-numeric cells, empty populations, and malformed
  // specs short-circuit upstream and return nullopt.
  return resolve_color_scale(rule, cell_value, ctx).has_value();
}

// ---------------------------------------------------------------------------
// DataBar — resolve `min`/`max` thresholds, then compute bar length and
// axis position.
//
// `length_pct` is the bar length expressed as a 0..100 percent of the
// cell width: cells at `min_threshold` produce `min_length_pct`; cells
// at `max_threshold` produce `max_length_pct`; cells outside clamp.
// `axis_position_pct` follows OOXML semantics: `Automatic` splits at the
// proportional negative offset, `Middle` pins to 50, `None` pins to 0.
// `is_negative` is set when the cell value is strictly negative so the
// host can flip the fill side.
// ---------------------------------------------------------------------------

constexpr double kAxisMid = 50.0;
constexpr double kAxisLeft = 0.0;
constexpr double kAxisRight = 100.0;

double automatic_axis_position(double threshold_min, double threshold_max) {
  // All non-negative → bar grows from the left edge.
  if (threshold_min >= 0.0) {
    return kAxisLeft;
  }
  // All non-positive → bar grows from the right edge.
  if (threshold_max <= 0.0) {
    return kAxisRight;
  }
  // Mixed sign: split proportionally so equal-magnitude positive and
  // negative bars meet at the same axis. Rare degenerate case
  // (threshold_min == 0 == threshold_max) is handled by the branches
  // above.
  const double negative_span = -threshold_min;
  const double total_span = negative_span + threshold_max;
  return (negative_span / total_span) * kAxisRight;
}

std::optional<DataBarRender> resolve_data_bar(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  if (!rule.data_bar.has_value()) {
    return std::nullopt;
  }
  const DataBarSpec& spec = *rule.data_bar;

  ColorScalePopulation fallback;
  const ColorScalePopulation* pop = numeric_population_for_cell(cell_value, ctx, fallback);
  if (pop == nullptr) {
    return std::nullopt;
  }

  auto threshold_min = resolve_cfvo(spec.min, *pop, ctx);
  auto threshold_max = resolve_cfvo(spec.max, *pop, ctx);
  if (!threshold_min.has_value() || !threshold_max.has_value()) {
    return std::nullopt;
  }
  // Degenerate threshold range collapses the bar — no meaningful length
  // can be produced. Fall through to nullopt so the caller can decide.
  if (*threshold_min == *threshold_max) {
    return std::nullopt;
  }

  const double cell = cell_value.as_number();
  const double range = *threshold_max - *threshold_min;
  const double raw_fraction = (cell - *threshold_min) / range;
  const double clamped_fraction = std::max(0.0, std::min(1.0, raw_fraction));
  const auto min_len = static_cast<double>(spec.min_length_pct);
  const auto max_len = static_cast<double>(spec.max_length_pct);

  DataBarRender render;
  render.length_pct = min_len + clamped_fraction * (max_len - min_len);
  render.is_negative = cell < 0.0;
  render.fill = render.is_negative ? spec.negative_fill : spec.fill;
  render.border = render.is_negative ? spec.negative_border : spec.border;
  render.gradient = spec.gradient;

  switch (spec.axis_position) {
    case DataBarAxisPosition::None:
      render.axis_position_pct = kAxisLeft;
      break;
    case DataBarAxisPosition::Middle:
      render.axis_position_pct = kAxisMid;
      break;
    case DataBarAxisPosition::Automatic:
      render.axis_position_pct = automatic_axis_position(*threshold_min, *threshold_max);
      break;
  }
  return render;
}

bool match_data_bar(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  return resolve_data_bar(rule, cell_value, ctx).has_value();
}

// ---------------------------------------------------------------------------
// IconSet — bucket the cell value across `N - 1` thresholds and assign
// an icon index in `[0, N - 1]` for an N-icon set.
//
// Each threshold's `gte` flag toggles `>=` vs. `>` at that boundary. The
// loop walks every threshold and bumps the index for each one the cell
// passes — assuming the OOXML reader populated thresholds in ascending
// order, which Excel always emits. `reverse` flips the index so the
// default-up direction can be inverted without re-sorting the colour /
// icon resources.
// ---------------------------------------------------------------------------

std::optional<IconRender> resolve_icon_set(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  if (!rule.icon_set.has_value()) {
    return std::nullopt;
  }
  const IconSetSpec& spec = *rule.icon_set;
  if (spec.thresholds.empty()) {
    return std::nullopt;  // A 1-icon "set" has no boundaries to match against.
  }

  ColorScalePopulation fallback;
  const ColorScalePopulation* pop = numeric_population_for_cell(cell_value, ctx, fallback);
  if (pop == nullptr) {
    return std::nullopt;
  }

  auto resolved = resolve_cfvo_list(spec.thresholds, *pop, ctx);
  if (!resolved.has_value()) {
    return std::nullopt;
  }

  const double cell = cell_value.as_number();
  std::uint8_t icon_index = 0;
  for (std::size_t i = 0; i < resolved->size(); ++i) {
    const bool above = spec.thresholds[i].gte ? (cell >= (*resolved)[i]) : (cell > (*resolved)[i]);
    if (above) {
      icon_index = static_cast<std::uint8_t>(i + 1);
    }
  }

  if (spec.reverse) {
    const auto bucket_count = static_cast<std::uint8_t>(resolved->size() + 1);
    icon_index = static_cast<std::uint8_t>(bucket_count - 1 - icon_index);
  }

  IconRender render;
  render.set_name = spec.name;
  render.icon_index = icon_index;
  return render;
}

bool match_icon_set(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  return resolve_icon_set(rule, cell_value, ctx).has_value();
}

}  // namespace

CFMatch make_match(const CFRule& rule) {
  CFMatch match;
  match.rule_id = rule.id;
  match.priority = rule.priority;
  // Value-only overload covers dxf-driven rules. Visual rule kinds
  // need the cell value and `CFEvalContext` to resolve their render
  // payload; callers should use the context-aware overload below.
  match.kind = CFMatchKind::DifferentialFormat;
  match.dxf_id = rule.dxf_id;
  return match;
}

namespace {

bool sqref_contains(const std::vector<CFCellRange>& sqref, CellAddress target) {
  for (const CFCellRange& range : sqref) {
    if (target.row >= range.first.row && target.row <= range.last.row && target.col >= range.first.col &&
        target.col <= range.last.col) {
      return true;
    }
  }
  return false;
}

CellAddress sqref_anchor(const std::vector<CFCellRange>& sqref) {
  // Excel authors CF formulas at the first cell of the first sqref
  // range; the shifter rebases relative refs from there.
  return sqref.empty() ? CellAddress{} : sqref.front().first;
}

}  // namespace

CFMatch make_match(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  CFMatch match;
  match.rule_id = rule.id;
  match.priority = rule.priority;
  if (rule.type == RuleType::ColorScale) {
    match.kind = CFMatchKind::ColorScale;
    match.resolved_fill_color = resolve_color_scale(rule, cell_value, ctx);
    return match;
  }
  if (rule.type == RuleType::DataBar) {
    match.kind = CFMatchKind::DataBar;
    match.data_bar_render = resolve_data_bar(rule, cell_value, ctx);
    return match;
  }
  if (rule.type == RuleType::IconSet) {
    match.kind = CFMatchKind::IconSet;
    match.icon_render = resolve_icon_set(rule, cell_value, ctx);
    return match;
  }
  // Every other kind is dxf-driven; the context-aware overload behaves
  // identically to the value-only one for those.
  match.kind = CFMatchKind::DifferentialFormat;
  match.dxf_id = rule.dxf_id;
  return match;
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
      return match_color_scale(rule, cell_value, ctx);
    case RuleType::DataBar:
      return match_data_bar(rule, cell_value, ctx);
    case RuleType::IconSet:
      return match_icon_set(rule, cell_value, ctx);
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

namespace {

// Per-block lazy population cache used by `evaluate_cf_for_range`.
// Indexed by block index inside `Sheet::conditional_formats()`. Slots
// stay empty until a rule that consumes the population first runs,
// then the slot is populated and reused across every cell in the range.
using PopulationCache = std::vector<std::optional<ColorScalePopulation>>;

// Whether `type` is a range-aware rule kind that benefits from a
// cached population. CellIs / Expression / TimePeriod / text-family /
// blanks-family / errors-family rules don't need it.
bool rule_uses_population(RuleType type) {
  switch (type) {
    case RuleType::ColorScale:
    case RuleType::DataBar:
    case RuleType::IconSet:
    case RuleType::AboveAverage:
    case RuleType::Top10:
      return true;
    default:
      return false;
  }
}

// Shared body for `evaluate_cf_at` and the viewport walker. When
// `cache` is non-null, the function lazily populates per-block slots
// the first time a rule needs the population and reuses them on
// subsequent calls. When null, every range-aware rule re-walks the
// sheet (the public single-cell behaviour).
std::vector<CFMatch> evaluate_cf_at_impl(const Sheet& sheet, CellAddress target, const CFHost& host,
                                         PopulationCache* cache) {
  std::vector<CFMatch> matches;
  if (host.arena == nullptr || host.registry == nullptr || host.eval_ctx == nullptr) {
    return matches;
  }

  // Collect (block_index, rule_index, priority) for every rule whose
  // sqref contains `target`. Indices instead of pointers keeps the
  // collection trivially copyable; sorting by priority is stable.
  struct Candidate {
    std::size_t block_index;
    std::size_t rule_index;
    std::int32_t priority;
  };
  std::vector<Candidate> candidates;

  const std::vector<ConditionalFormat>& blocks = sheet.conditional_formats();
  for (std::size_t block_idx = 0; block_idx < blocks.size(); ++block_idx) {
    if (!sqref_contains(blocks[block_idx].sqref, target)) {
      continue;
    }
    for (std::size_t rule_idx = 0; rule_idx < blocks[block_idx].rules.size(); ++rule_idx) {
      candidates.push_back({block_idx, rule_idx, blocks[block_idx].rules[rule_idx].priority});
    }
  }
  std::stable_sort(candidates.begin(), candidates.end(),
                   [](const Candidate& lhs, const Candidate& rhs) { return lhs.priority < rhs.priority; });

  const Value cell_value = sheet.resolve_cell_value(target.row, target.col);
  for (const Candidate& candidate : candidates) {
    const ConditionalFormat& block = blocks[candidate.block_index];
    const CFRule& rule = block.rules[candidate.rule_index];

    CFEvalContext ctx;
    ctx.anchor = sqref_anchor(block.sqref);
    ctx.target = target;
    ctx.arena = host.arena;
    ctx.registry = host.registry;
    ctx.eval_ctx = host.eval_ctx;
    ctx.today_serial = host.today_serial;
    ctx.sqref = &block.sqref;

    // Lazy-populate the cache slot for this block when the rule needs
    // a numeric population. Non-cache callers (`cache == nullptr`)
    // keep `ctx.cached_population` null and the helpers gather on
    // demand.
    if (cache != nullptr && rule_uses_population(rule.type)) {
      std::optional<ColorScalePopulation>& slot = (*cache)[candidate.block_index];
      if (!slot.has_value()) {
        slot = gather_population(block.sqref, sheet);
      }
      ctx.cached_population = &*slot;
    }

    if (!match_rule(rule, cell_value, ctx)) {
      continue;
    }
    matches.push_back(make_match(rule, cell_value, ctx));
    if (rule.stop_if_true) {
      break;
    }
  }
  return matches;
}

}  // namespace

std::vector<CFMatch> evaluate_cf_at(const Sheet& sheet, CellAddress target, const CFHost& host) {
  return evaluate_cf_at_impl(sheet, target, host, /*cache=*/nullptr);
}

std::vector<CFRangeCellMatches> evaluate_cf_for_range(const Sheet& sheet, CFCellRange range, const CFHost& host) {
  std::vector<CFRangeCellMatches> results;
  if (host.arena == nullptr || host.registry == nullptr || host.eval_ctx == nullptr) {
    return results;
  }
  // One cache for the whole range. Slots are sized to match the
  // sheet's block count; each slot is populated lazily the first time
  // a range-aware rule needs it, then reused across every subsequent
  // cell in the same block.
  PopulationCache cache(sheet.conditional_formats().size());
  // Iterate row-major over the inclusive range. The boundary check is
  // `<=` so a single-cell range (first == last) still produces one
  // visit. Callers who need to evaluate one cell should prefer the
  // direct `evaluate_cf_at`; this helper exists for the viewport case.
  for (std::uint32_t row = range.first.row; row <= range.last.row; ++row) {
    for (std::uint32_t col = range.first.col; col <= range.last.col; ++col) {
      CellAddress cell{};
      cell.row = row;
      cell.col = col;
      std::vector<CFMatch> matches = evaluate_cf_at_impl(sheet, cell, host, &cache);
      if (matches.empty()) {
        continue;
      }
      CFRangeCellMatches entry;
      entry.cell = cell;
      entry.matches = std::move(matches);
      results.push_back(std::move(entry));
    }
  }
  return results;
}

}  // namespace formulon::cf
