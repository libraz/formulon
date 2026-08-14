//
// Implementation of `EvalContext::resolve_ref`. The contract — in particular
// the full mapping from `Reference` shape to returned `Value` — lives in the
// header Doxygen; this file only executes it.

#include "eval/eval_context.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
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
#include "sheet_name.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
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

// An ordinary cell reference is one scalar coordinate even when its fresh
// recursive formula evaluation produces an Array (for example, SEQUENCE at
// a committed spill anchor). The explicit SpillRef (`A1#`) path bypasses
// resolve_ref and is therefore the only way to request the whole array.
// Collapse here, at the shared reference boundary, so direct refs, normal
// ranges and 3-D ranges all agree while stale formula caches still get fresh
// recursive resolution. Preserve the empty-array error contract.
Value scalarize_formula_ref(Value value) {
  if (!value.is_array()) {
    return value;
  }
  if (value.as_array_rows() == 0U || value.as_array_cols() == 0U) {
    return Value::error(ErrorCode::Value);
  }
  return value.as_array_cells()[0];
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

  // Same-sheet mutable references are dispatched through SpillCommitter and
  // therefore already return the anchor scalar. A read-only context must
  // scalarize an ordinary formula reference itself. A cross-sheet reference
  // from a mutable context remains an uncommitted Array for the outer formula
  // owner; this preserves the existing no-cross-sheet-spill contract.
  const bool mutable_target = mutable_sheet_ != nullptr && prefix.target_sheet == mutable_sheet_;
  const bool scalarize_result = mutable_sheet_ == nullptr || mutable_target;

  // Formula cell. If no recursive state is bound, fall back to the
  // non-recursive behaviour so the two overloads agree.
  if (state_ == nullptr) {
    return scalarize_result ? scalarize_formula_ref(prefix.cell->cached_value) : prefix.cell->cached_value;
  }

  if (const Value* memo = state_->lookup_memo(prefix.target_sheet, prefix.row, prefix.col); memo != nullptr) {
    return scalarize_result ? scalarize_formula_ref(*memo) : *memo;
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
      return scalarize_result ? scalarize_formula_ref(prefix.cell->cached_value) : prefix.cell->cached_value;
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

    // Route every mutable recursive result through SpillCommitter. Arrays are
    // committed as usual; scalar/error results clear any previously committed
    // or blocked state at the referenced formula anchor. Cross-sheet results
    // remain read-only because spill commits across sheets are out of scope.
    if (mutable_target) {
      result = child_ctx.dispatch_array_result(result);
    }
    if (scalarize_result) {
      // A read-only recursive context cannot commit a dynamic-array result,
      // so an ordinary cell reference still exposes one scalar coordinate.
      result = scalarize_formula_ref(result);
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
  return SpillCommitter(anchor_sheet, formula_row_, formula_col_, spill_release_callback_, spill_release_user_data_)
      .commit(std::move(v));
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
  // shape. When both endpoints carry a qualifier they must agree under
  // locale-independent Unicode simple case folding or the range is `#REF!`.
  std::string_view effective_sheet_name;
  if (!lhs.sheet.empty() && !rhs.sheet.empty()) {
    if (!sheet_names::equal(lhs.sheet, rhs.sheet)) {
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
      const std::optional<Sheet::PopulatedExtent> extent =
          target_sheet->populated_extent(0U, c_min, Sheet::kMaxRows - 1U, c_max);
      if (!extent.has_value()) {
        set_shape(out_rows, out_cols, 0, 0);
        return std::vector<Value>{};
      }
      r_min = 0;
      r_max = extent->last_row;
    } else {
      r_min = std::min(lhs.row, rhs.row);
      r_max = std::max(lhs.row, rhs.row);
      if (r_max >= Sheet::kMaxRows) {
        return ErrorCode::Ref;
      }
      const std::optional<Sheet::PopulatedExtent> extent =
          target_sheet->populated_extent(r_min, 0U, r_max, Sheet::kMaxCols - 1U);
      if (!extent.has_value()) {
        set_shape(out_rows, out_cols, 0, 0);
        return std::vector<Value>{};
      }
      c_min = 0;
      c_max = extent->last_col;
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
  if (total > kMaxRangeExpansionCells) {
    return ErrorCode::Calc;
  }
  out.reserve(static_cast<std::size_t>(total));
  // Bulk read: one lock acquisition for the whole rectangle instead of two
  // per cell. The sheet was already resolved above, and every coordinate is
  // in bounds, so this reproduces what a per-cell `resolve_ref` would return
  // for literal, phantom and absent cells. Formula cells come back as their
  // cached value and are re-resolved below.
  std::vector<std::size_t> formula_indices;
  target_sheet->read_range(r_min, r_max, c_min, c_max, out, formula_indices);
  const std::uint32_t width = c_max - c_min + 1U;
  for (const std::size_t index : formula_indices) {
    parser::Reference cell_ref{};
    // Propagate the effective sheet qualifier so `resolve_ref` routes
    // through the same `resolve_target_sheet` logic — this keeps the
    // `(sheet, row, col)` cycle key for the correct target sheet even
    // for ranges where only LHS was qualified in the source.
    cell_ref.sheet = effective_sheet_name;
    cell_ref.row = r_min + static_cast<std::uint32_t>(index / width);
    cell_ref.col = c_min + static_cast<std::uint32_t>(index % width);
    // Per-cell error Values (e.g. #DIV/0! from a formula cell, #REF!
    // from a cycle caught by EvalState) are pushed through unchanged so
    // the dispatcher can honour `propagate_errors`. The shared resolver
    // scalarizes ordinary formula references, including 3-D callers.
    out[index] = resolve_ref(cell_ref, arena, registry);
  }
  // Mark only the copied range values. The source sheet/cache remains plain
  // Blank; range-aware consumers still observe the Blank kind, while the
  // terminal grid boundary can project these raw-grid cells to zero.
  for (Value& value : out) {
    if (value.is_blank()) {
      value = Value::blank(BlankGridProjection::kReferenceGridZero);
    }
  }
  set_shape(out_rows, out_cols, r_max - r_min + 1U, c_max - c_min + 1U);
  return out;
}

}  // namespace eval
}  // namespace formulon
