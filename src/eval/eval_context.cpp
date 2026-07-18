// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of `EvalContext::resolve_ref`. The contract — in particular
// the full mapping from `Reference` shape to returned `Value` — lives in the
// header Doxygen; this file only executes it.

#include "eval/eval_context.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "cell.h"
#include "eval/eval_state.h"
#include "eval/formula_text_utils.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/spill_committer.h"
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
const Sheet* resolve_target_sheet(std::string_view sheet_name, const Sheet* current_sheet, const Workbook* workbook,
                                  ErrorCode* out_err) {
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

// True when a stored cell carries no observable content (neither a formula
// nor a non-blank cached value). Mirrors the used-range predicate the
// print-layout pagination pass applies so both agree on a sheet's extent.
bool cell_is_blank(const Cell& cell) {
  return cell.formula_text.empty() && cell.cached_value.is_blank();
}

// Largest 0-based row index holding a non-blank cell within columns
// [col_lo, col_hi] (inclusive). Returns false when no such cell exists, in
// which case `*out_max_row` is left untouched. Used to bound whole-column
// expansion so `SUM(A:A)` walks only up to the column's populated extent
// instead of all `Sheet::kMaxRows` rows.
bool max_row_in_cols(const Sheet& sheet, std::uint32_t col_lo, std::uint32_t col_hi, std::uint32_t* out_max_row) {
  bool any = false;
  std::uint32_t max_row = 0;
  for (const auto& [row_index, cells] : sheet.rows()) {
    const std::size_t upper = std::min<std::size_t>(cells.size(), static_cast<std::size_t>(col_hi) + 1U);
    for (std::size_t c = col_lo; c < upper; ++c) {
      if (cell_is_blank(cells[c])) {
        continue;
      }
      if (!any || row_index > max_row) {
        max_row = row_index;
        any = true;
      }
      break;
    }
  }
  if (!any) {
    return false;
  }
  *out_max_row = max_row;
  return true;
}

// Largest 0-based column index holding a non-blank cell within rows
// [row_lo, row_hi] (inclusive). Returns false when no such cell exists.
// Bounds whole-row expansion so `COUNTA(1:1)` walks only up to the row's
// populated extent instead of all `Sheet::kMaxCols` columns.
bool max_col_in_rows(const Sheet& sheet, std::uint32_t row_lo, std::uint32_t row_hi, std::uint32_t* out_max_col) {
  bool any = false;
  std::uint32_t max_col = 0;
  for (const auto& [row_index, cells] : sheet.rows()) {
    if (row_index < row_lo || row_index > row_hi) {
      continue;
    }
    for (std::size_t c = cells.size(); c-- > 0;) {
      if (cell_is_blank(cells[c])) {
        continue;
      }
      const auto col_index = static_cast<std::uint32_t>(c);
      if (!any || col_index > max_col) {
        max_col = col_index;
        any = true;
      }
      break;
    }
  }
  if (!any) {
    return false;
  }
  *out_max_col = max_col;
  return true;
}

// Writes the computed shape into the optional out-params, tolerating null.
void set_shape(std::uint32_t* out_rows, std::uint32_t* out_cols, std::uint32_t rows, std::uint32_t cols) {
  if (out_rows != nullptr) {
    *out_rows = rows;
  }
  if (out_cols != nullptr) {
    *out_cols = cols;
  }
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
ResolvePrefix resolve_prefix(const Sheet* current_sheet, const Workbook* workbook, const parser::Reference& ref) {
  ResolvePrefix out;
  if (current_sheet == nullptr) {
    out.value = Value::error(ErrorCode::Name);
    return out;
  }
  // Cross-sheet resolution: without a workbook or a known target sheet,
  // `#REF!`. With a match, the target becomes the current sheet for the
  // remainder of the preamble.
  ErrorCode sheet_err = ErrorCode::Ref;
  const Sheet* target = resolve_target_sheet(ref.sheet, current_sheet, workbook, &sheet_err);
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
  // Phantom-aware read: a cell that is either absent, or stored as a literal,
  // is fed through `resolve_cell_value` so that phantom cells of a committed
  // spill region are visible to cross-cell references. The formula branch
  // below still requires the stored cell because we need `formula_text`.
  if (cell == nullptr || cell->formula_text.empty()) {
    out.value = target->resolve_cell_value(ref.row, ref.col);
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

Value EvalContext::resolve_ref(const parser::Reference& ref, Arena& arena, const FunctionRegistry& registry) const {
  const ResolvePrefix prefix = resolve_prefix(current_sheet_, workbook_, ref);
  if (prefix.kind == ResolvePrefix::Kind::Terminal) {
    return prefix.value;
  }

  // Formula cell. If no recursive state is bound, fall back to the
  // non-recursive behaviour so the two overloads agree.
  if (state_ == nullptr) {
    return prefix.cell->cached_value;
  }

  if (const Value* memo = state_->lookup_memo(prefix.target_sheet, prefix.row, prefix.col); memo != nullptr) {
    return *memo;
  }

  if (!state_->push_cell(prefix.target_sheet, prefix.row, prefix.col)) {
    // Direct or indirect cycle (possibly spanning sheets). With Excel's
    // "Enable iterative calculation" option on, a back-edge inside a cycle
    // resolves to the cell's last computed value (the iterative-calc driver
    // in `evaluate()` re-runs the anchor formula until it reaches a fixed
    // point, seeding each pass from the prior value). With iterative calc
    // off, `#REF!` is the closest Excel-observable error — Excel itself
    // shows 0 plus a circular-reference warning banner, which Formulon has
    // no surface for yet.
    if (workbook_ != nullptr && workbook_->iterative_options().enabled) {
      return prefix.cell->cached_value;
    }
    return Value::error(ErrorCode::Ref);
  }

  // Parse `formula_text` in the caller's evaluation arena. Reusing a single
  // arena keeps text payloads readable after the recursive call returns.
  // `Cell::formula_text` is stored verbatim by `Sheet::set_cell_formula`,
  // so it may carry the leading `=` from external readers or hand-written
  // callers. Strip it through the shared helper so this resolver agrees
  // bit-for-bit with `recalc_engine.cpp` / `scheduler.cpp`.
  parser::AstNode* root = parser::parse_strict(strip_formula_prefix(prefix.cell->formula_text), arena);

  Value result = Value::blank();
  if (root == nullptr) {
    // Hard parse failure, or a valid prefix trailed by unparseable tokens.
    // We refuse to resolve a recovered prefix as the referenced cell's
    // value; #NAME? matches the recursive-resolver failure contract.
    result = Value::error(ErrorCode::Name);
  } else {
    // Anchor the recursive evaluation at the target cell so ROW() / COLUMN()
    // inside the referenced formula return the referenced cell's coordinates
    // rather than inheriting the caller's anchor. The new context inherits
    // `mutable_sheet_` (regular member, copied by `with_formula_cell`).
    const EvalContext child_ctx = this->with_formula_cell(prefix.row, prefix.col);
    result = evaluate(*root, arena, registry, child_ctx);

    // If the recursive formula produced an Array AND the target cell lives
    // on the sheet we are authorised to mutate, commit the spill now and
    // memoise the resulting scalar (anchor cell value, or #SPILL! on
    // collision). Cross-sheet recursive arrays are intentionally left
    // un-spilled here: spill commits across sheets are out of scope and
    // would mutate a sheet the caller did not opt in to.
    if (mutable_sheet_ != nullptr && prefix.target_sheet == mutable_sheet_ && result.is_array()) {
      result = child_ctx.dispatch_array_result(result);
    }
  }

  state_->pop_cell(prefix.target_sheet, prefix.row, prefix.col);
  state_->memoize(prefix.target_sheet, prefix.row, prefix.col, result);
  return result;
}

Value EvalContext::dispatch_array_result(Value v) const {
  // Compatibility wrapper: the actual spill logic now lives on
  // `SpillCommitter` so non-`EvalContext` callers (and future call sites
  // that only need write authority) can opt in without inflating a full
  // evaluation context. An unbound `mutable_sheet_` or a missing formula-
  // cell anchor produces an inactive committer, which passes the array
  // through unchanged — preserving the historical behaviour exactly.
  Sheet* anchor_sheet = (mutable_sheet_ != nullptr && has_formula_cell()) ? mutable_sheet_ : nullptr;
  return SpillCommitter(anchor_sheet, formula_row_, formula_col_).commit(std::move(v));
}

Expected<std::vector<Value>, ErrorCode> EvalContext::expand_range(const parser::Reference& lhs,
                                                                  const parser::Reference& rhs, Arena& arena,
                                                                  const FunctionRegistry& registry,
                                                                  std::uint32_t* out_rows,
                                                                  std::uint32_t* out_cols) const {
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
  const Sheet* target_sheet = resolve_target_sheet(effective_sheet_name, current_sheet_, workbook_, &sheet_err);
  if (target_sheet == nullptr) {
    return sheet_err;
  }

  std::uint32_t r_min = 0;
  std::uint32_t r_max = 0;
  std::uint32_t c_min = 0;
  std::uint32_t c_max = 0;
  if (lhs.is_full_col || lhs.is_full_row || rhs.is_full_col || rhs.is_full_row) {
    // Whole-column (`A:A` / `A:C`) and whole-row (`1:1` / `1:3`) references
    // are expanded against the target sheet's used range: the unbounded
    // axis is clamped to the sheet's populated extent so the expansion
    // never physically walks all `kMaxRows` / `kMaxCols` cells, while the
    // bounded axis keeps its natural origin (row 0 for a column, column 0
    // for a row) so positional consumers (INDEX / VLOOKUP column offsets)
    // see the reference's true top-left. A sheet with no content in range
    // yields an empty expansion, so `SUM(A:A)` on an empty sheet is 0
    // without any per-row work.
    //
    // Only same-axis whole references compose (`A:C`, `1:3`); a mixed
    // whole-column / whole-row pair has no bounded rectangle and degrades
    // to #VALUE!, matching the pre-existing scalar degradation.
    const bool col_range = lhs.is_full_col && rhs.is_full_col;
    const bool row_range = lhs.is_full_row && rhs.is_full_row;
    if (!col_range && !row_range) {
      return ErrorCode::Value;
    }
    if (col_range) {
      c_min = std::min(lhs.col, rhs.col);
      c_max = std::max(lhs.col, rhs.col);
      if (c_max >= Sheet::kMaxCols) {
        return ErrorCode::Ref;
      }
      std::uint32_t max_row = 0;
      if (!max_row_in_cols(*target_sheet, c_min, c_max, &max_row)) {
        set_shape(out_rows, out_cols, 0, 0);
        return std::vector<Value>{};
      }
      r_min = 0;
      r_max = max_row;
    } else {
      r_min = std::min(lhs.row, rhs.row);
      r_max = std::max(lhs.row, rhs.row);
      if (r_max >= Sheet::kMaxRows) {
        return ErrorCode::Ref;
      }
      std::uint32_t max_col = 0;
      if (!max_col_in_rows(*target_sheet, r_min, r_max, &max_col)) {
        set_shape(out_rows, out_cols, 0, 0);
        return std::vector<Value>{};
      }
      c_min = 0;
      c_max = max_col;
    }
  } else {
    if (lhs.row >= Sheet::kMaxRows || lhs.col >= Sheet::kMaxCols || rhs.row >= Sheet::kMaxRows ||
        rhs.col >= Sheet::kMaxCols) {
      return ErrorCode::Ref;
    }
    // Normalise endpoint ordering: A3:A1 describes the same rectangle as
    // A1:A3.
    r_min = std::min(lhs.row, rhs.row);
    r_max = std::max(lhs.row, rhs.row);
    c_min = std::min(lhs.col, rhs.col);
    c_max = std::max(lhs.col, rhs.col);
  }

  // Accepted divergence: callers such as SUM / AVERAGE coerce every
  // expanded Value via `coerce_to_number`, so a range cell holding TRUE
  // contributes 1.0 (not 0) and a text cell surfaces `#VALUE!` instead of
  // being silently skipped. Excel's range-vs-direct semantic split is a
  // future pass; flattening here keeps the dispatcher agnostic.
  std::vector<Value> out;
  const std::uint64_t total =
      static_cast<std::uint64_t>(r_max - r_min + 1) * static_cast<std::uint64_t>(c_max - c_min + 1);
  // Refuse pathological rectangles before `reserve()`. Excel's logical
  // grid is 1,048,576 rows * 16,384 cols ~= 1.7e10 cells, which would
  // overflow `std::size_t` on a 32-bit WASM build (`size_t` is 32-bit
  // there) and silently truncate the reservation, leaving the inner
  // push-back loop to repeatedly re-allocate / OOM. Capping at 10M cells
  // is still well above any realistic aggregator footprint and rejects
  // `A1:XFD1048576`-shaped inputs uniformly across native and WASM.
  constexpr std::uint64_t kMaxRangeExpansionCells = 10'000'000ULL;
  if (total > kMaxRangeExpansionCells) {
    return ErrorCode::Calc;
  }
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
  set_shape(out_rows, out_cols, r_max - r_min + 1U, c_max - c_min + 1U);
  return out;
}

}  // namespace eval
}  // namespace formulon
