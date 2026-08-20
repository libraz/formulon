//
// Implementation of `EvalContext::resolve_ref`. The contract — in particular
// the full mapping from `Reference` shape to returned `Value` — lives in the
// header Doxygen; this file only executes it.

#include "eval/eval_context.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string_view>
#include <vector>

#include "eval/declared_rect.h"
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

// Common short-circuit checks shared by both overloads: unbound context,
// cross-sheet lookup, full-column / full-row, out-of-bounds. On success the
// resolved owning sheet is returned (equal to `current_sheet` for
// unqualified refs, or the workbook-looked-up sheet for qualified ones); on
// any short circuit the function returns `nullptr` and writes the Excel
// error sentinel the caller must return into `*out_value`.
const Sheet* resolve_ref_target(const Sheet* current_sheet, const Workbook* workbook, const parser::Reference& ref,
                                Value* out_value) {
  if (current_sheet == nullptr) {
    *out_value = Value::error(ErrorCode::Name);
    return nullptr;
  }
  // Cross-sheet resolution: without a workbook or a known target sheet,
  // `#REF!`. With a match, the target becomes the current sheet for the
  // remainder of the preamble.
  ErrorCode sheet_err = ErrorCode::Ref;
  const Sheet* target = resolve_target_sheet(ref.sheet, current_sheet, workbook, &sheet_err);
  if (target == nullptr) {
    *out_value = Value::error(sheet_err);
    return nullptr;
  }
  if (ref.is_full_col || ref.is_full_row) {
    // Whole-column / whole-row references are ranges; in scalar context they
    // degrade to #VALUE! until array evaluation lands.
    *out_value = Value::error(ErrorCode::Value);
    return nullptr;
  }
  if (ref.row >= Sheet::kMaxRows || ref.col >= Sheet::kMaxCols) {
    *out_value = Value::error(ErrorCode::Ref);
    return nullptr;
  }
  return target;
}

// Effective sheet qualifier for a rectangle.
//
// The parser retains the `:` operator's qualifier on the LHS in practice, so
// `Sheet2!A1:B2` parses as RangeOp(Ref{sheet=Sheet2}, Ref{sheet=""}) — the RHS
// must inherit. Defensive symmetry is kept for the opposite shape. When both
// endpoints carry a qualifier they must agree under locale-independent Unicode
// simple case folding or the range is `#REF!`. An empty result means the
// context's own sheet.
Expected<std::string_view, ErrorCode> effective_range_sheet(const parser::Reference& lhs,
                                                            const parser::Reference& rhs) {
  if (!lhs.sheet.empty() && !rhs.sheet.empty()) {
    if (!sheet_names::equal(lhs.sheet, rhs.sheet)) {
      return ErrorCode::Ref;
    }
    return lhs.sheet;
  }
  if (!lhs.sheet.empty()) {
    return lhs.sheet;
  }
  return rhs.sheet;
}

// Re-homes a Text payload into `arena`.
//
// A `Sheet::CellRead` owns its bytes only for as long as the snapshot
// object lives, and the snapshot is a local of `resolve_ref`. The arena is
// the lifetime the overload already promises its callers, so anything
// derived from a snapshot has to be copied into it before it is returned.
// Non-Text values carry no pointer and pass through untouched.
Value adopt_into_arena(Arena& arena, Value value) {
  if (!value.is_text()) {
    return value;
  }
  return Value::text(arena.intern(value.as_text()));
}

}  // namespace

Value EvalContext::resolve_ref(const parser::Reference& ref) const {
  Value short_circuit = Value::blank();
  const Sheet* target = resolve_ref_target(current_sheet_, workbook_, ref, &short_circuit);
  if (target == nullptr) {
    return short_circuit;
  }
  // Non-recursive path: a formula cell yields whatever is currently cached
  // (typically blank), which is exactly what the spill-aware read returns
  // for it — a formula cell is either an ordinary cell or a spill anchor,
  // and an anchor's `cached_value` tracks its region's first cell. Reading
  // through the sheet keeps the whole lookup inside one lock acquisition
  // and hands no `Cell*` back to this layer.
  return target->resolve_cell_value(ref.row, ref.col);
}

Value EvalContext::resolve_ref(const parser::Reference& ref, Arena& arena, const FunctionRegistry& registry) const {
  Value short_circuit = Value::blank();
  const Sheet* target = resolve_ref_target(current_sheet_, workbook_, ref, &short_circuit);
  if (target == nullptr) {
    return short_circuit;
  }

  // One lock acquisition covers "is this a formula cell" and both fields a
  // formula cell is read for. Holding a `Cell*` past the lock instead would
  // let a concurrent `set_cell_cached_value` rewrite `formula_text` and
  // `cached_value` — and free the Text buffer `cached_value` points at —
  // underneath the reads below.
  Sheet::CellRead read;
  target->read_formula_cell(ref.row, ref.col, read);
  if (!read.is_formula()) {
    // Phantom-aware read: an absent or literal cell resolves through the
    // spill table so phantom cells of a committed region are visible to
    // cross-cell references.
    return adopt_into_arena(arena, read.value());
  }

  // Same-sheet mutable references are dispatched through SpillCommitter and
  // therefore already return the anchor scalar. A read-only context must
  // scalarize an ordinary formula reference itself. A cross-sheet reference
  // from a mutable context remains an uncommitted Array for the outer formula
  // owner; this preserves the existing no-cross-sheet-spill contract.
  const bool mutable_target = mutable_sheet_ != nullptr && target == mutable_sheet_;
  const bool scalarize_result = mutable_sheet_ == nullptr || mutable_target;

  // Formula cell. If no recursive state is bound, fall back to the
  // non-recursive behaviour so the two overloads agree.
  if (state_ == nullptr) {
    const Value cached = adopt_into_arena(arena, read.value());
    return scalarize_result ? scalarize_formula_ref(cached) : cached;
  }

  if (const Value* memo = state_->lookup_memo(target, ref.row, ref.col); memo != nullptr) {
    return scalarize_result ? scalarize_formula_ref(*memo) : *memo;
  }

  if (!state_->push_cell(target, ref.row, ref.col)) {
    // Direct or indirect cycle (possibly spanning sheets). With Excel's
    // "Enable iterative calculation" option on, a back-edge inside a cycle
    // resolves to the cell's last computed value (the iterative-calc driver
    // in `evaluate()` re-runs the anchor formula until it reaches a fixed
    // point, seeding each pass from the prior value). With iterative calc
    // off, `#REF!` is the closest Excel-observable error — Excel itself
    // shows 0 plus a circular-reference warning banner, which Formulon has
    // no surface for yet.
    if (workbook_ != nullptr && workbook_->iterative_options().enabled) {
      const Value cached = adopt_into_arena(arena, read.value());
      return scalarize_result ? scalarize_formula_ref(cached) : cached;
    }
    return Value::error(ErrorCode::Ref);
  }

  // Parse `formula_text` in the caller's evaluation arena. Reusing a single
  // arena keeps text payloads readable after the recursive call returns.
  // `Cell::formula_text` is stored verbatim by `Sheet::set_cell_formula`,
  // so it may carry the leading `=` from external readers or hand-written
  // callers. Strip it through the shared helper so this resolver agrees
  // bit-for-bit with `recalc_engine.cpp` / `scheduler.cpp`.
  //
  // The source is interned first: the AST keeps views into it, a string
  // literal in the formula surfaces in `result` as a view into the same
  // bytes, and both outlive the snapshot the text was copied out of.
  const std::string_view source = arena.intern(strip_formula_prefix(read.formula_text()));
  parser::AstNode* root = parser::parse_strict(source, arena);

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
    const EvalContext child_ctx = this->with_formula_cell(ref.row, ref.col);
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

  state_->pop_cell(target, ref.row, ref.col);
  state_->memoize(target, ref.row, ref.col, result);
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

  auto effective_sheet = effective_range_sheet(lhs, rhs);
  if (!effective_sheet) {
    return effective_sheet.error();
  }
  const std::string_view effective_sheet_name = effective_sheet.value();

  ErrorCode sheet_err = ErrorCode::Ref;
  const Sheet* target_sheet = resolve_target_sheet(effective_sheet_name, current_sheet_, workbook_, &sheet_err);
  if (target_sheet == nullptr) {
    return sheet_err;
  }

  // The rectangle the endpoints declare, derived where every other consumer
  // of a static reference derives it. Endpoint ordering is normalised there,
  // so A3:A1 describes the same rectangle as A1:A3.
  const auto declared = declared_rect(lhs, rhs);
  if (!declared) {
    return declared.error();
  }
  std::uint32_t r_min = declared.value().row_first;
  std::uint32_t r_max = declared.value().row_last;
  std::uint32_t c_min = declared.value().col_first;
  std::uint32_t c_max = declared.value().col_last;
  if (declared.value().whole_axis) {
    // Enumeration-only narrowing. A whole-column (`A:A` / `A:C`) or
    // whole-row (`1:1` / `1:3`) reference declares a whole grid axis; the
    // values worth walking end at the target sheet's populated extent, so
    // the unbounded axis is clamped to it and the expansion never physically
    // touches all `kMaxRows` / `kMaxCols` cells. `SUM(A:A)` on an empty
    // sheet is 0 without any per-row work.
    //
    // The bounded axis keeps its natural origin (row 0 for a column, column
    // 0 for a row) so positional consumers (INDEX / VLOOKUP column offsets)
    // see the reference's true top-left.
    //
    // The clamp is invisible to a consumer of the values — the cells it
    // drops are empty — and must stay invisible to everything else. It is
    // not the reference's shape, and `declared_range_rect` exists so that no
    // caller needing a shape has to reach for it.
    const std::optional<Sheet::PopulatedExtent> extent = target_sheet->populated_extent(r_min, c_min, r_max, c_max);
    if (!extent.has_value()) {
      set_shape(out_rows, out_cols, 0, 0);
      return std::vector<Value>{};
    }
    if (declared.value().rows() == Sheet::kMaxRows) {
      r_max = extent->last_row;
    } else {
      c_max = extent->last_col;
    }
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
  // cached value and are re-resolved below. Text payloads land in `arena`,
  // the same lifetime the scalar path promises.
  std::vector<std::size_t> formula_indices;
  target_sheet->read_range(r_min, r_max, c_min, c_max, arena, out, formula_indices);
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

Expected<DeclaredRect, ErrorCode> EvalContext::declared_range_rect(const parser::Reference& lhs,
                                                                   const parser::Reference& rhs) const {
  if (current_sheet_ == nullptr) {
    return ErrorCode::Name;
  }
  auto effective_sheet = effective_range_sheet(lhs, rhs);
  if (!effective_sheet) {
    return effective_sheet.error();
  }
  // The rectangle itself is sheet-independent, but a qualifier naming a
  // sheet the workbook does not hold is still `#REF!` — a shape must not be
  // reported for a reference that cannot be resolved at all.
  ErrorCode sheet_err = ErrorCode::Ref;
  if (resolve_target_sheet(effective_sheet.value(), current_sheet_, workbook_, &sheet_err) == nullptr) {
    return sheet_err;
  }
  return declared_rect(lhs, rhs);
}

}  // namespace eval
}  // namespace formulon
