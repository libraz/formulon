#include "validation_eval.h"

#include <algorithm>
#include <cmath>
#include <cstdio>
#include <cstdlib>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_evaluator.h"
#include "cf/cf_helpers.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/resource_budget.h"
#include "utils/utf8_length.h"
#include "workbook.h"

namespace formulon {
namespace {

// DataValidation::type / ::op encodings (see sheet.h).
constexpr std::uint8_t kTypeNone = 0;
constexpr std::uint8_t kTypeWhole = 1;
constexpr std::uint8_t kTypeDecimal = 2;
constexpr std::uint8_t kTypeList = 3;
constexpr std::uint8_t kTypeDate = 4;
constexpr std::uint8_t kTypeTime = 5;
constexpr std::uint8_t kTypeTextLength = 6;
constexpr std::uint8_t kTypeCustom = 7;

constexpr std::uint8_t kOpBetween = 0;
constexpr std::uint8_t kOpNotBetween = 1;
constexpr std::uint8_t kOpEqual = 2;
constexpr std::uint8_t kOpNotEqual = 3;
constexpr std::uint8_t kOpGreaterThan = 4;
constexpr std::uint8_t kOpLessThan = 5;
constexpr std::uint8_t kOpGreaterThanOrEqual = 6;
constexpr std::uint8_t kOpLessThanOrEqual = 7;

bool contains(const DataValidation& dv, std::uint32_t row, std::uint32_t col) {
  return std::any_of(dv.ranges.begin(), dv.ranges.end(), [&](const MergeRange& r) {
    return row >= r.first_row && row <= r.last_row && col >= r.first_col && col <= r.last_col;
  });
}

// Top-left corner of the bounding box of the rule's sqref.
CellAddress anchor_of(const DataValidation& dv) {
  CellAddress a{0U, 0U};
  bool first = true;
  for (const MergeRange& r : dv.ranges) {
    a.row = first ? r.first_row : std::min(a.row, r.first_row);
    a.col = first ? r.first_col : std::min(a.col, r.first_col);
    first = false;
  }
  return a;
}

/// Evaluation environment shared by every formula of one call.
struct Evaluator {
  const Workbook& wb;
  const Sheet& sheet;
  CellAddress anchor;
  CellAddress target;
  Arena arena{/*initial_chunk_bytes=*/4096, kMaxEvalArenaBytes};
  eval::EvalState state;
  eval::EvalContext eval_ctx;
  const eval::FunctionRegistry* registry;

  Evaluator(const Workbook& w, const Sheet& s, CellAddress anc, CellAddress tgt, const eval::FunctionRegistry* reg)
      : wb(w), sheet(s), anchor(anc), target(tgt), eval_ctx(w, s, state), registry(reg) {}

  Value run(const std::string& source) {
    cf::CFEvalContext ctx;
    ctx.anchor = anchor;
    ctx.target = target;
    ctx.arena = &arena;
    ctx.registry = registry;
    ctx.eval_ctx = &eval_ctx;
    return cf::helpers::parse_shift_evaluate(source, ctx);
  }

  /// Evaluates `source` to a scalar; a 1x1 array collapses to its element.
  Value scalar(const std::string& source) {
    Value v = run(source);
    if (v.is_array() && v.as_array_rows() == 1U && v.as_array_cols() == 1U) {
      return v.as_array()->cells[0];
    }
    return v;
  }
};

std::optional<double> as_bound(const Value& v) {
  if (v.is_number()) {
    return v.as_number();
  }
  if (v.is_blank()) {
    return 0.0;
  }
  return std::nullopt;
}

bool compare(std::uint8_t op, double v, double lo, double hi) {
  switch (op) {
    case kOpBetween:
      return v >= lo && v <= hi;
    case kOpNotBetween:
      return v < lo || v > hi;
    case kOpEqual:
      return v == lo;
    case kOpNotEqual:
      return v != lo;
    case kOpGreaterThan:
      return v > lo;
    case kOpLessThan:
      return v < lo;
    case kOpGreaterThanOrEqual:
      return v >= lo;
    case kOpLessThanOrEqual:
      return v <= lo;
    default:
      return false;
  }
}

/// Resolves formula1 (and formula2 for the between operators) and applies the operator.
bool check_operator(Evaluator& ev, const DataValidation& dv, double v) {
  const std::optional<double> lo = as_bound(ev.scalar(dv.formula1));
  if (!lo.has_value()) {
    return false;
  }
  double hi = 0.0;
  if (dv.op == kOpBetween || dv.op == kOpNotBetween) {
    const std::optional<double> h = as_bound(ev.scalar(dv.formula2));
    if (!h.has_value()) {
      return false;
    }
    hi = *h;
  }
  return compare(dv.op, v, *lo, hi);
}

bool check_numeric(Evaluator& ev, const DataValidation& dv, const Value& proposed) {
  double v = 0.0;
  if (proposed.is_number()) {
    v = proposed.as_number();
  } else if (!proposed.is_blank()) {
    return false;
  }
  if (!std::isfinite(v)) {
    return false;
  }
  if (dv.type == kTypeWhole && v != std::floor(v)) {
    return false;
  }
  return check_operator(ev, dv, v);
}

bool check_text_length(Evaluator& ev, const DataValidation& dv, const Value& proposed) {
  auto text = eval::coerce_to_text(proposed);
  if (!text) {
    return false;
  }
  return check_operator(ev, dv, static_cast<double>(utf16_units_in(text.value())));
}

/// Parses an inline list body (`"a,b,c"`, with `""` as an escaped quote) into its items.
std::vector<std::string> inline_items(const std::string& formula) {
  std::string body;
  for (std::size_t i = 1; i + 1 < formula.size(); ++i) {
    body.push_back(formula[i]);
    if (formula[i] == '"' && formula[i + 1] == '"') {
      ++i;
    }
  }
  std::vector<std::string> items;
  std::size_t start = 0;
  for (;;) {
    const std::size_t comma = body.find(',', start);
    items.push_back(body.substr(start, comma == std::string::npos ? std::string::npos : comma - start));
    if (comma == std::string::npos) {
      break;
    }
    start = comma + 1;
  }
  return items;
}

bool is_inline_list(const std::string& f) {
  return f.size() >= 2U && f.front() == '"' && f.back() == '"';
}

std::optional<double> parse_item_number(const std::string& item) {
  if (item.empty()) {
    return std::nullopt;
  }
  char* end = nullptr;
  const double d = std::strtod(item.c_str(), &end);
  if (end != item.c_str() + item.size() || !std::isfinite(d)) {
    return std::nullopt;
  }
  return d;
}

bool inline_list_matches(const std::vector<std::string>& items, const Value& proposed) {
  for (const std::string& item : items) {
    if (proposed.is_text()) {
      if (proposed.as_text() == item) {
        return true;
      }
    } else if (proposed.is_number()) {
      const std::optional<double> n = parse_item_number(item);
      if (n.has_value() && *n == proposed.as_number()) {
        return true;
      }
    }
  }
  return false;
}

/// Renders `proposed` as a formula literal; nullopt for kinds a list cannot contain.
std::optional<std::string> literal_of(const Value& proposed) {
  if (proposed.is_number()) {
    char buf[40];
    std::snprintf(buf, sizeof(buf), "%.17g", proposed.as_number());
    return std::string(buf);
  }
  if (proposed.is_boolean()) {
    return std::string(proposed.as_boolean() ? "TRUE" : "FALSE");
  }
  if (proposed.is_text()) {
    // MATCH reads `*`, `?` and `~` in text as wildcards; a literal list entry must not.
    std::string lit = "\"";
    for (const char c : proposed.as_text()) {
      if (c == '*' || c == '?' || c == '~') {
        lit.push_back('~');
      }
      lit.push_back(c);
      if (c == '"') {
        lit.push_back('"');
      }
    }
    lit.push_back('"');
    return lit;
  }
  return std::nullopt;
}

bool check_list(Evaluator& ev, const DataValidation& dv, const Value& proposed) {
  if (is_inline_list(dv.formula1)) {
    return inline_list_matches(inline_items(dv.formula1), proposed);
  }
  // A range source is tested with an exact MATCH: a spilled range read at the target cell
  // would be blocked by the cells below it, and MATCH gives Excel's case-insensitive,
  // type-strict equality.
  const std::optional<std::string> lit = literal_of(proposed);
  if (!lit.has_value()) {
    return false;
  }
  const Value found = ev.scalar("ISNUMBER(MATCH(" + *lit + "," + dv.formula1 + ",0))");
  return found.is_boolean() && found.as_boolean();
}

bool check_custom(Evaluator& ev, const DataValidation& dv) {
  const Value v = ev.scalar(dv.formula1);
  if (v.is_boolean()) {
    return v.as_boolean();
  }
  return v.is_number() && v.as_number() != 0.0;
}

}  // namespace

Expected<ValidationOutcome, Error> validate_value(const Workbook& wb, const Sheet& sheet, std::uint32_t row,
                                                  std::uint32_t col, const Value& proposed,
                                                  const ValidationEvalDeps& deps) {
  using Result = Expected<ValidationOutcome, Error>;
  if (!Sheet::rect_in_grid(row, col, row, col)) {
    return Result::Err(make_error(FormulonErrorCode::kInvalidArgument, "validate_value: cell out of grid"));
  }
  ValidationOutcome out;
  const std::vector<DataValidation>& rules = sheet.validations();
  for (std::size_t i = 0; i < rules.size(); ++i) {
    if (!contains(rules[i], row, col)) {
      continue;
    }
    const DataValidation& dv = rules[i];
    out.has_rule = true;
    out.rule_index = static_cast<std::uint32_t>(i);
    out.error_style = dv.error_style;
    if (dv.type == kTypeNone || (proposed.is_blank() && dv.allow_blank)) {
      out.valid = true;
      return Result::Ok(out);
    }
    const eval::FunctionRegistry* registry = deps.registry != nullptr ? deps.registry : &eval::default_registry();
    Evaluator ev(wb, sheet, anchor_of(dv), CellAddress{row, col}, registry);
    switch (dv.type) {
      case kTypeWhole:
      case kTypeDecimal:
      case kTypeDate:
      case kTypeTime:
        out.valid = check_numeric(ev, dv, proposed);
        break;
      case kTypeTextLength:
        out.valid = check_text_length(ev, dv, proposed);
        break;
      case kTypeList:
        out.valid = check_list(ev, dv, proposed);
        break;
      case kTypeCustom:
        out.valid = check_custom(ev, dv);
        break;
      default:
        out.valid = false;
        break;
    }
    return Result::Ok(out);
  }
  return Result::Ok(out);
}

}  // namespace formulon
