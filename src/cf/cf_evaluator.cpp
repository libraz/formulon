// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the CF evaluator. See cf_evaluator.h for the
// scoped contract and the staged-PR roadmap.

#include "cf/cf_evaluator.h"

#include <algorithm>
#include <cstdlib>
#include <optional>
#include <string>
#include <string_view>

#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "value.h"

namespace formulon::cf {

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
    case RuleType::ContainsText:
    case RuleType::NotContainsText:
    case RuleType::BeginsWith:
    case RuleType::EndsWith:
    case RuleType::TimePeriod:
    case RuleType::DuplicateValues:
    case RuleType::UniqueValues:
      return false;
  }
  return false;
}

CFMatch make_match(const CFRule& rule) {
  CFMatch match;
  match.rule_id = rule.id;
  match.priority = rule.priority;
  // The skeleton landing covers only dxf-driven rule types; later PRs
  // widen this dispatch to assign `kind = ColorScale / DataBar /
  // IconSet` and populate the corresponding render payload.
  match.kind = CFMatchKind::DifferentialFormat;
  match.dxf_id = rule.dxf_id;
  return match;
}

}  // namespace formulon::cf
