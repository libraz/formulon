//
// Implementation of the ad-hoc, side-effect-free formula evaluation
// drivers declared in `adhoc_eval.h`.

#include "eval/adhoc_eval.h"

#include <cstdint>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/formula_text_utils.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "parser/ast.h"
#include "parser/ast_shift.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

namespace {

// Reduces a possibly-array `Value` to a single scalar by taking the
// top-left cell of an array result. See the header for why this is neither
// implicit intersection nor spilling. A degenerate empty array (which
// producers should never emit) surfaces as `#VALUE!`.
Value reduce_to_scalar(Value v) {
  if (!v.is_array()) {
    return v;
  }
  if (v.as_array_rows() == 0 || v.as_array_cols() == 0) {
    return Value::error(ErrorCode::Value);
  }
  return v.as_array_cells()[0];
}

// Builds the read-only evaluation context shared by both drivers. The
// deliberate omission of `with_mutable_sheet` is the purity guarantee:
// without it `EvalContext::dispatch_array_result` is inert and recursive
// `resolve_ref` never commits a spill, so the workbook is never mutated.
EvalContext make_readonly_context(const Workbook& workbook, const Sheet& sheet, EvalState& state, std::uint32_t row,
                                  std::uint32_t col) {
  return EvalContext(workbook, sheet, state)
      .with_excel_profile(workbook.excel_profile())
      .with_date1904(workbook.date1904())
      .with_pinned_now(workbook.pinned_now())
      .with_formula_cell(row, col);
}

// Parses and evaluates `formula` at `(row, col)` under a read-only context,
// returning the raw `Value` WITHOUT the array-to-scalar reduction. Shared by
// both the scalar-reducing and array-preserving public drivers so their
// parse / anchor / purity behaviour stays identical. Parser failure surfaces
// as `#NAME?`, matching `EvalContext::resolve_ref`. The `EvalState` is local:
// the returned `Value`'s text / array payloads borrow `arena` (which the
// caller owns), never the transient state.
Value parse_and_evaluate(const Workbook& workbook, const Sheet& sheet, std::uint32_t row, std::uint32_t col,
                         std::string_view formula, Arena& arena, const FunctionRegistry& registry) {
  // Strip a leading '=' so both "=A1+1" and "A1+1" parse identically,
  // matching the recalc path's use of `strip_formula_prefix`.
  const std::string_view src = strip_formula_prefix(formula);

  parser::AstNode* root = parser::parse_strict(src, arena);
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }

  EvalState state;
  const EvalContext ctx = make_readonly_context(workbook, sheet, state, row, col);
  return evaluate(*root, arena, registry, ctx);
}

}  // namespace

Value evaluate_formula_text(const Workbook& workbook, const Sheet& sheet, std::uint32_t row, std::uint32_t col,
                            std::string_view formula, Arena& arena, const FunctionRegistry& registry) {
  return reduce_to_scalar(parse_and_evaluate(workbook, sheet, row, col, formula, arena, registry));
}

Value evaluate_formula_text_array(const Workbook& workbook, const Sheet& sheet, std::uint32_t row, std::uint32_t col,
                                  std::string_view formula, Arena& arena, const FunctionRegistry& registry) {
  return parse_and_evaluate(workbook, sheet, row, col, formula, arena, registry);
}

bool evaluate_cf_formula(const Workbook& workbook, const Sheet& sheet, std::uint32_t row, std::uint32_t col,
                         std::uint32_t anchor_row, std::uint32_t anchor_col, std::string_view formula, Arena& arena,
                         const FunctionRegistry& registry) {
  const std::string_view src = strip_formula_prefix(formula);

  parser::AstNode* root = parser::parse_strict(src, arena);
  if (root == nullptr) {
    // A malformed rule formula does not fire (coerces to false).
    return false;
  }

  // Relative refs in a CF rule formula are authored relative to the
  // applied range's top-left (`anchor`); shift them to the target cell so
  // they resolve exactly as Excel does for `(row, col)`.
  const std::int32_t row_delta = static_cast<std::int32_t>(row) - static_cast<std::int32_t>(anchor_row);
  const std::int32_t col_delta = static_cast<std::int32_t>(col) - static_cast<std::int32_t>(anchor_col);
  const parser::AstNode* shifted = parser::shift_relative_refs(*root, arena, row_delta, col_delta);
  if (shifted == nullptr) {
    return false;
  }

  EvalState state;
  const EvalContext ctx = make_readonly_context(workbook, sheet, state, row, col);
  const Value result = reduce_to_scalar(evaluate(*shifted, arena, registry, ctx));
  return coerce_cf_predicate(result);
}

bool coerce_cf_predicate(const Value& v) {
  if (v.is_boolean()) {
    return v.as_boolean();
  }
  if (v.is_number()) {
    return v.as_number() != 0.0;
  }
  // Error, blank, text, array, lambda: the rule does not fire.
  return false;
}

}  // namespace eval
}  // namespace formulon
