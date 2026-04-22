// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of `EvalContext::resolve_ref`. The contract — in particular
// the full mapping from `Reference` shape to returned `Value` — lives in the
// header Doxygen; this file only executes it.

#include "eval/eval_context.h"

#include <algorithm>
#include <cstdint>
#include <string_view>
#include <vector>

#include "cell.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Resolves the target sheet for a reference given a (possibly empty) sheet
// name, the bound current sheet, and the bound workbook.
//
// On success, returns a non-null `Sheet*`. On failure — the ref is
// qualified but no workbook is bound, or the named sheet is not present —
// returns `nullptr` and writes the error code (`Ref`) into `*out_err`.
//
// Assumes `current_sheet` is non-null when `sheet_name` is empty; callers
// must short-circuit the unbound-context case before invoking this helper.
const Sheet* resolve_target_sheet(std::string_view sheet_name,
                                  const Sheet* current_sheet,
                                  const Workbook* workbook, ErrorCode* out_err) {
  if (sheet_name.empty()) {
    return current_sheet;
  }
  if (workbook == nullptr) {
    *out_err = ErrorCode::Ref;
    return nullptr;
  }
  const Sheet* target = workbook->sheet_by_name(sheet_name);
  if (target == nullptr) {
    *out_err = ErrorCode::Ref;
    return nullptr;
  }
  return target;
}

// Result of the shared preamble executed by both `resolve_ref` overloads.
// The `value` field carries either the short-circuit Value (literal cell,
// absent cell, or Excel error sentinel) or a valid pointer to the formula
// cell whose `formula_text` needs evaluation. `target_sheet` is the resolved
// owning sheet for formula cells (equal to `current_sheet` for unqualified
// refs, or the workbook-looked-up sheet for qualified refs).
struct ResolvePrefix {
  enum class Kind : std::uint8_t {
    /// Preamble produced a definitive Value (error, blank, or literal cache)
    /// and the caller should return it unchanged.
    Terminal,
    /// Preamble found a formula cell; the caller decides whether to recurse.
    Formula,
  };
  Kind kind = Kind::Terminal;
  // Populated when kind == Terminal.
  Value value = Value::blank();
  // Populated when kind == Formula.
  const Cell* cell = nullptr;
  const Sheet* target_sheet = nullptr;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
};

// Common short-circuit checks shared by both overloads: unbound context,
// cross-sheet lookup, full-column / full-row, out-of-bounds, absent cell,
// literal cell. A formula-cell hit is returned as `Kind::Formula` so each
// caller can decide whether (and how) to evaluate the formula.
ResolvePrefix resolve_prefix(const Sheet* current_sheet,
                             const Workbook* workbook,
                             const parser::Reference& ref) {
  ResolvePrefix out;
  if (current_sheet == nullptr) {
    out.value = Value::error(ErrorCode::Name);
    return out;
  }
  // Cross-sheet resolution: without a workbook or a known target sheet,
  // `#REF!`. With a match, the target becomes the current sheet for the
  // remainder of the preamble.
  ErrorCode sheet_err = ErrorCode::Ref;
  const Sheet* target =
      resolve_target_sheet(ref.sheet, current_sheet, workbook, &sheet_err);
  if (target == nullptr) {
    out.value = Value::error(sheet_err);
    return out;
  }
  if (ref.is_full_col || ref.is_full_row) {
    // Whole-column / whole-row references are ranges; in scalar context they
    // degrade to #VALUE! until array evaluation lands.
    out.value = Value::error(ErrorCode::Value);
    return out;
  }
  if (ref.row >= Sheet::kMaxRows || ref.col >= Sheet::kMaxCols) {
    out.value = Value::error(ErrorCode::Ref);
    return out;
  }
  const Cell* cell = target->cell_at(ref.row, ref.col);
  if (cell == nullptr) {
    out.value = Value::blank();
    return out;
  }
  if (cell->formula_text.empty()) {
    // Literal cell — hand back its cached value verbatim.
    out.value = cell->cached_value;
    return out;
  }
  out.kind = ResolvePrefix::Kind::Formula;
  out.cell = cell;
  out.target_sheet = target;
  out.row = ref.row;
  out.col = ref.col;
  return out;
}

}  // namespace

Value EvalContext::resolve_ref(const parser::Reference& ref) const {
  const ResolvePrefix prefix = resolve_prefix(current_sheet_, workbook_, ref);
  if (prefix.kind == ResolvePrefix::Kind::Terminal) {
    return prefix.value;
  }
  // Formula cell, non-recursive path: mirror the historical behaviour by
  // returning whatever is currently cached (typically blank).
  return prefix.cell->cached_value;
}

Value EvalContext::resolve_ref(const parser::Reference& ref, Arena& arena,
                               const FunctionRegistry& registry) const {
  const ResolvePrefix prefix = resolve_prefix(current_sheet_, workbook_, ref);
  if (prefix.kind == ResolvePrefix::Kind::Terminal) {
    return prefix.value;
  }

  // Formula cell. If no recursive state is bound, fall back to the
  // non-recursive behaviour so the two overloads agree.
  if (state_ == nullptr) {
    return prefix.cell->cached_value;
  }

  if (const Value* memo =
          state_->lookup_memo(prefix.target_sheet, prefix.row, prefix.col);
      memo != nullptr) {
    return *memo;
  }

  if (!state_->push_cell(prefix.target_sheet, prefix.row, prefix.col)) {
    // Direct or indirect cycle (possibly spanning sheets). Excel shows
    // 0 + a warning banner for iterative-calc-disabled cycles, but
    // Formulon has no banner surface yet; `#REF!` is the closest
    // Excel-observable error. Iterative calc and SCC pre-detection are a
    // later phase.
    return Value::error(ErrorCode::Ref);
  }

  // Parse `formula_text` in the caller's evaluation arena. Reusing a single
  // arena keeps text payloads readable after the recursive call returns.
  parser::Parser parser(prefix.cell->formula_text, arena);
  parser::AstNode* root = parser.parse();

  Value result = Value::blank();
  if (root == nullptr) {
    // Parser failed beyond recovery (e.g. empty input or arena exhaustion).
    // Malformed-but-recoverable inputs already become #NAME? via the
    // evaluator's handling of `ErrorPlaceholder`, so both paths agree.
    result = Value::error(ErrorCode::Name);
  } else {
    result = evaluate(*root, arena, registry, *this);
  }

  state_->pop_cell(prefix.target_sheet, prefix.row, prefix.col);
  state_->memoize(prefix.target_sheet, prefix.row, prefix.col, result);
  return result;
}

Expected<std::vector<Value>, ErrorCode> EvalContext::expand_range(
    const parser::Reference& lhs, const parser::Reference& rhs, Arena& arena,
    const FunctionRegistry& registry) const {
  if (current_sheet_ == nullptr) {
    return ErrorCode::Name;
  }

  // Infer the effective sheet qualifier for the rectangle. The parser
  // retains the `:` operator's qualifier on the LHS in practice, so
  // `Sheet2!A1:B2` parses as RangeOp(Ref{sheet=Sheet2}, Ref{sheet=""}) —
  // the RHS must inherit. Defensive symmetry is kept for the opposite
  // shape. When both endpoints carry a qualifier they must agree
  // (case-insensitively) or the range is `#REF!`.
  std::string_view effective_sheet_name;
  if (!lhs.sheet.empty() && !rhs.sheet.empty()) {
    if (!strings::case_insensitive_eq(lhs.sheet, rhs.sheet)) {
      return ErrorCode::Ref;
    }
    effective_sheet_name = lhs.sheet;
  } else if (!lhs.sheet.empty()) {
    effective_sheet_name = lhs.sheet;
  } else if (!rhs.sheet.empty()) {
    effective_sheet_name = rhs.sheet;
  }

  ErrorCode sheet_err = ErrorCode::Ref;
  const Sheet* target_sheet = resolve_target_sheet(
      effective_sheet_name, current_sheet_, workbook_, &sheet_err);
  if (target_sheet == nullptr) {
    return sheet_err;
  }

  if (lhs.is_full_col || lhs.is_full_row || rhs.is_full_col || rhs.is_full_row) {
    // Whole-column / whole-row ranges are deferred until array evaluation
    // lands; in scalar context they degrade to #VALUE!.
    return ErrorCode::Value;
  }
  if (lhs.row >= Sheet::kMaxRows || lhs.col >= Sheet::kMaxCols ||
      rhs.row >= Sheet::kMaxRows || rhs.col >= Sheet::kMaxCols) {
    return ErrorCode::Ref;
  }

  // Normalise endpoint ordering: A3:A1 describes the same rectangle as
  // A1:A3.
  const std::uint32_t r_min = std::min(lhs.row, rhs.row);
  const std::uint32_t r_max = std::max(lhs.row, rhs.row);
  const std::uint32_t c_min = std::min(lhs.col, rhs.col);
  const std::uint32_t c_max = std::max(lhs.col, rhs.col);

  // Accepted divergence: callers such as SUM / AVERAGE coerce every
  // expanded Value via `coerce_to_number`, so a range cell holding TRUE
  // contributes 1.0 (not 0) and a text cell surfaces `#VALUE!` instead of
  // being silently skipped. Excel's range-vs-direct semantic split is a
  // future pass; flattening here keeps the dispatcher agnostic.
  std::vector<Value> out;
  const std::uint64_t total = static_cast<std::uint64_t>(r_max - r_min + 1) *
                              static_cast<std::uint64_t>(c_max - c_min + 1);
  out.reserve(static_cast<std::size_t>(total));
  for (std::uint32_t r = r_min; r <= r_max; ++r) {
    for (std::uint32_t c = c_min; c <= c_max; ++c) {
      parser::Reference cell_ref{};
      // Propagate the effective sheet qualifier so `resolve_ref` routes
      // through the same `resolve_target_sheet` logic — this keeps the
      // `(sheet, row, col)` cycle key for the correct target sheet even
      // for ranges where only LHS was qualified in the source.
      cell_ref.sheet = effective_sheet_name;
      cell_ref.row = r;
      cell_ref.col = c;
      // Per-cell error Values (e.g. #DIV/0! from a formula cell, #REF!
      // from a cycle caught by EvalState) are pushed through unchanged so
      // the dispatcher can honour `propagate_errors`.
      out.push_back(resolve_ref(cell_ref, arena, registry));
    }
  }
  return out;
}

}  // namespace eval
}  // namespace formulon
