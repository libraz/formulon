// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the CF evaluator's shared helpers. See cf_helpers.h
// for the contract.

#include "cf/cf_helpers.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_evaluator.h"
#include "cf/cf_types.h"
#include "eval/coerce.h"
#include "eval/date_time.h"
#include "eval/eval_context.h"
#include "eval/scalar_ops.h"
#include "eval/tree_walker.h"
#include "parser/ast.h"
#include "parser/ast_shift.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/rect_iterator.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon::cf::helpers {

// ---------------------------------------------------------------------------
// Literal-operand helpers.
// ---------------------------------------------------------------------------

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

  // Numeric literal — consumed in full or rejected. Route through the
  // evaluator's shared `strtod_full` scanner so the cellIs literal parser
  // and the engine's text-to-number coercion accept exactly the same syntax.
  double parsed = 0.0;
  if (eval::strtod_full(source, &parsed)) {
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

std::optional<int> compare_cell_to_literal(const Value& cell, const LiteralOperand& operand) {
  if (cell.is_error() || cell.is_blank()) {
    return std::nullopt;
  }

  // Build a `Value` view of the operand and route the comparison through the
  // evaluator's shared `compare_values` so the cellIs rule sees exactly the
  // ordering / case-folding the engine applies to the `=` / `<` / `>`
  // operators. `LiteralOperand` owns its text payload, so the non-owning
  // `Value::text` view is valid for the duration of this call.
  switch (operand.kind) {
    case LiteralOperand::Kind::Text: {
      // A text operand only meaningfully orders against a text cell; a
      // numeric / boolean cell leaves the rule inactive (Excel does not
      // cross-coerce here).
      if (!cell.is_text()) {
        return std::nullopt;
      }
      bool unordered = false;
      return eval::compare_values(cell, Value::text(operand.text_value), &unordered);
    }
    case LiteralOperand::Kind::Number:
    case LiteralOperand::Kind::Bool: {
      // A numeric / boolean operand orders against numeric or boolean cells;
      // a text cell leaves the rule inactive. Both sides are lifted onto the
      // numeric line (Bool -> 1/0) so a Bool cell and a Number operand (or
      // vice versa) order by value — Excel compares cellIs numeric operands
      // by magnitude, not by `compare_values`'s cross-kind Number < Bool
      // ranking.
      if (!cell.is_number() && !cell.is_boolean()) {
        return std::nullopt;
      }
      const double cell_num = cell.is_number() ? cell.as_number() : (cell.as_boolean() ? 1.0 : 0.0);
      const double lit_num = literal_as_number(operand);
      if (cell_num < lit_num) {
        return -1;
      }
      if (cell_num > lit_num) {
        return 1;
      }
      return 0;
    }
  }
  return std::nullopt;
}

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

std::optional<LiteralOperand> cell_is_operand(const std::string& source, const CFEvalContext& ctx) {
  auto literal = parse_literal(source);
  if (literal.has_value()) {
    return literal;
  }
  const Value evaluated = parse_shift_evaluate(source, ctx);
  return value_to_operand(evaluated);
}

// ---------------------------------------------------------------------------
// Substring / prefix / suffix helpers used by the text-rule family.
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Population gathering / caching.
// ---------------------------------------------------------------------------

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

std::size_t count_matches_in_sqref(const Value& target, const std::vector<CFCellRange>& sqref, const Sheet& sheet) {
  std::size_t count = 0;
  for (const CFCellRange& range : sqref) {
    for (auto [row, col] : utils::RectRange(range.first.row, range.first.col, range.last.row, range.last.col)) {
      const Value cell = sheet.resolve_cell_value(row, col);
      if (cf_values_equal(target, cell)) {
        ++count;
      }
    }
  }
  return count;
}

std::vector<double> collect_numeric_values(const std::vector<CFCellRange>& sqref, const Sheet& sheet) {
  std::vector<double> values;
  for (const CFCellRange& range : sqref) {
    for (auto [row, col] : utils::RectRange(range.first.row, range.first.col, range.last.row, range.last.col)) {
      const Value cell = sheet.resolve_cell_value(row, col);
      if (cell.is_number()) {
        values.push_back(cell.as_number());
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

// ---------------------------------------------------------------------------
// Numeric helpers.
// ---------------------------------------------------------------------------

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
  // Shared scanner with the engine's text-to-number coercion so CFVO
  // thresholds parse identically to `=`/`<`/`>` numeric operands.
  double parsed = 0.0;
  if (eval::strtod_full(source, &parsed)) {
    return parsed;
  }
  return std::nullopt;
}

// ---------------------------------------------------------------------------
// Date helpers.
// ---------------------------------------------------------------------------

namespace {

constexpr int kMonthsPerYear = 12;
constexpr int kDaysPerWeek = 7;

}  // namespace

int weekday_sunday_one(double serial_floor) {
  const eval::date_time::YMD ymd = eval::date_time::ymd_from_serial(serial_floor);
  const std::int64_t days = eval::date_time::days_from_civil(ymd.y, ymd.m, ymd.d);
  // 1970-01-01 was a Thursday → Excel weekday 5. Adjust so days = 0
  // maps to 5, then take mod 7 and shift to the 1..7 range.
  const std::int64_t adjusted = ((days % kDaysPerWeek) + kDaysPerWeek + 4) % kDaysPerWeek;
  return static_cast<int>(adjusted) + 1;
}

double sunday_of_week(double serial_floor) {
  const int weekday = weekday_sunday_one(serial_floor);
  return serial_floor - (weekday - 1);
}

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

}  // namespace formulon::cf::helpers
