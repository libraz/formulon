// Copyright 2026 libraz. Licensed under the MIT License.
//
// `EvalContext` is the abstraction through which the tree-walk evaluator
// resolves cell references. It carries the sheet-binding that anchors local
// (unqualified) A1 references — without it, `NodeKind::Ref` cannot be
// resolved and the evaluator falls back to `#NAME?`.
//
// A context may also carry an optional `EvalState`. When bound, references
// to formula cells are recursively parsed and evaluated on demand, with
// per-call memoisation and cycle detection (cycles surface as `#REF!`).
// Without an `EvalState`, formula cells return their cached value verbatim
// (which is typically blank, because nothing populates it in that mode).
//
// When a context is constructed with a `Workbook` (the three-arg form),
// qualified references (`Reference::sheet` non-empty) are looked up
// case-insensitively in the workbook and resolved against the matching
// sheet. Cycle detection and memoisation live on `(sheet, row, col)` via
// `EvalState`, so cross-sheet cycles are caught. A context bound only to a
// `Sheet` (two-arg form) still resolves qualified references to `#REF!`,
// because there is no workbook to query.

#ifndef FORMULON_EVAL_EVAL_CONTEXT_H_
#define FORMULON_EVAL_EVAL_CONTEXT_H_

#include <cstdint>
#include <vector>

#include "parser/reference.h"
#include "sheet.h"
#include "utils/expected.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {

class Arena;
class Workbook;

namespace eval {

class EvalState;
class FunctionRegistry;
class NameEnv;

/// Evaluator-side view of the data a formula needs to resolve cell
/// references.
///
/// An `EvalContext` is one of three shapes:
///   * Unbound (default-constructed): every reference resolves to `#NAME?`,
///     preserving the evaluator's behaviour before reference resolution
///     was wired in.
///   * Bound to a single `Sheet`: local A1 references resolve against it;
///     qualified references still resolve to `#REF!` because there is no
///     workbook to look up the target sheet.
///   * Bound to a `Workbook` + current `Sheet` + `EvalState`: local refs
///     resolve against the current sheet; qualified refs are looked up
///     case-insensitively in the workbook. Cross-sheet cycles are caught
///     via the `(sheet, row, col)` key in `EvalState`.
///
/// A context is a lightweight, non-owning view — the referenced `Sheet`,
/// `Workbook`, and `EvalState` must outlive every evaluator invocation
/// that uses the context.
class EvalContext {
 public:
  /// Builds an unbound context. `resolve_ref` resolves every reference to
  /// `#NAME?`, preserving the evaluator's behaviour before reference
  /// resolution was wired in.
  EvalContext() = default;

  /// Builds a context bound to `sheet` as the current sheet, without a
  /// recursive evaluation state. Formula cells return their cached value
  /// verbatim; no recursion happens. Qualified refs (`Reference::sheet`
  /// non-empty) resolve to `#REF!` because there is no workbook bound.
  explicit EvalContext(const Sheet& sheet) noexcept : current_sheet_(&sheet) {}

  /// Builds a context bound to `sheet` and a recursive evaluation `state`.
  ///
  /// When resolving a reference to a formula cell the 3-arg `resolve_ref`
  /// overload will parse and evaluate `formula_text` on demand, memoising
  /// the result into `state`. Direct or indirect cycles surface as
  /// `#REF!`. The sheet is NOT mutated by this process: `cell->cached_value`
  /// is left alone, and all formula-result caching lives in `state` for the
  /// duration of a single `evaluate()` call. Qualified refs resolve to
  /// `#REF!` (no workbook bound).
  EvalContext(const Sheet& sheet, EvalState& state) noexcept
      : current_sheet_(&sheet), state_(&state) {}

  /// Evaluating + cross-sheet–aware context. Unqualified refs resolve
  /// against `current_sheet`; qualified refs (`Reference::sheet` non-empty)
  /// are looked up in `workbook` case-insensitively. Cycle detection and
  /// memoisation operate on `(sheet, row, col)` via `state`, so cross-sheet
  /// cycles are caught.
  ///
  /// `current_sheet` MUST be a sheet owned by `workbook` (no check is
  /// enforced). `workbook`, `current_sheet`, and `state` must outlive the
  /// context.
  EvalContext(const Workbook& workbook, const Sheet& current_sheet,
              EvalState& state) noexcept
      : current_sheet_(&current_sheet), state_(&state), workbook_(&workbook) {}

  /// Resolves an A1 reference to the cell's cached `Value` (non-recursive).
  ///
  /// The mapping from `Reference` states to returned values:
  ///
  /// | Condition                                          | Result           |
  /// |----------------------------------------------------|------------------|
  /// | Context is unbound (`current_sheet() == nullptr`)  | `#NAME?`         |
  /// | `ref.sheet` non-empty and no workbook bound        | `#REF!`          |
  /// | `ref.sheet` non-empty and missing from workbook    | `#REF!`          |
  /// | `ref.is_full_col \|\| ref.is_full_row`             | `#VALUE!`        |
  /// | `ref.row >= Sheet::kMaxRows \|\|                   |                  |
  /// |  ref.col >= Sheet::kMaxCols`                       | `#REF!`          |
  /// | Cell is absent from storage                        | `Value::blank()` |
  /// | Cell exists                                        | cell cached_value|
  ///
  /// This overload never recurses: formula cells return `cached_value`
  /// verbatim even when `state()` is bound. Use the 3-arg overload for
  /// recursive evaluation.
  Value resolve_ref(const parser::Reference& ref) const;

  /// Resolves an A1 reference, recursively evaluating formula cells when
  /// `state()` is bound.
  ///
  /// Error / short-circuit cases (cross-sheet lookup failure, whole-column
  /// / whole-row, out-of-bounds, cell absent, literal cell) behave exactly
  /// like the 1-arg overload. The recursive path applies only when a
  /// formula cell is found AND `state()` is non-null:
  ///
  ///   1. If the result has already been memoised on `state_` for the
  ///      resolved target sheet, return it.
  ///   2. Try to push the cell's `(target_sheet, row, col)` address onto
  ///      the in-progress stack. On duplicate (a direct or indirect cycle,
  ///      including across sheets) return `#REF!`.
  ///   3. Parse `formula_text` in `arena`. A null root — parser failure
  ///      beyond recovery — is surfaced as `#NAME?`. Panic-mode recovery
  ///      otherwise substitutes `ErrorPlaceholder` nodes, which the
  ///      evaluator itself turns into `#NAME?` during evaluation.
  ///   4. Recursively evaluate the AST with `registry` and this context.
  ///   5. Pop the stack frame and memoise the result on `state_`.
  ///
  /// The Sheet is NOT mutated; `cell->cached_value` is left alone. Text
  /// values returned from recursion reference storage in `arena`, so the
  /// caller's evaluation arena must outlive the returned `Value`.
  Value resolve_ref(const parser::Reference& ref, Arena& arena,
                    const FunctionRegistry& registry) const;

  /// Expands the rectangle `[lhs : rhs]` into a flat list of cell values in
  /// row-major order (top-left to bottom-right, row by row). Endpoint
  /// ordering is normalised: `A3:A1` and `A1:A3` yield the same expansion.
  ///
  /// Each cell is resolved via `resolve_ref(cell_ref, arena, registry)` so
  /// formula cells are recursed into (subject to `EvalState` cycle detection
  /// when bound) and cell absence is reported as `Value::blank()`. Returning
  /// an error `Value` inside the vector is legal - the dispatcher propagates
  /// it per `propagate_errors`.
  ///
  /// Sheet-qualifier handling:
  ///   * Both endpoints unqualified → expands over the current sheet.
  ///   * Only `lhs` qualified (the common shape parsed from e.g.
  ///     `Sheet2!A1:B2`, where the `:` operator retains the qualifier on
  ///     the left) → expands over `lhs.sheet`.
  ///   * Only `rhs` qualified (defensive — parser does not emit this
  ///     shape in practice) → expands over `rhs.sheet`.
  ///   * Both qualified with mismatching names → `#REF!`.
  ///   * Qualified but no workbook bound → `#REF!`.
  ///   * Qualified but sheet not found in workbook → `#REF!`.
  ///
  /// Short-circuit error mapping (first match wins):
  ///
  /// | Condition                                                 | Result   |
  /// |-----------------------------------------------------------|----------|
  /// | Context is unbound (`current_sheet() == nullptr`)         | `#NAME?` |
  /// | Mismatched cross-sheet endpoints                          | `#REF!`  |
  /// | Qualified ref with no workbook bound                      | `#REF!`  |
  /// | Qualified ref whose target sheet is missing               | `#REF!`  |
  /// | Either endpoint has `is_full_col` or `is_full_row`        | `#VALUE!`|
  /// | Either endpoint has row/col >= `Sheet::kMax*`             | `#REF!`  |
  /// | Otherwise                                                 | vector   |
  Expected<std::vector<Value>, ErrorCode> expand_range(
      const parser::Reference& lhs, const parser::Reference& rhs,
      Arena& arena, const FunctionRegistry& registry) const;

  /// Returns the sheet this context is bound to, or `nullptr` when the
  /// context was default-constructed.
  const Sheet* current_sheet() const noexcept { return current_sheet_; }

  /// Returns the recursive-evaluation state this context is bound to, or
  /// `nullptr` when no state was supplied. Useful for tests that want to
  /// inspect the memoisation map after an evaluation.
  EvalState* state() const noexcept { return state_; }

  /// Returns the workbook this context is bound to, or `nullptr` when the
  /// context was built without one. Exposed so consumers (tests, debug
  /// printers) can observe whether cross-sheet resolution is available.
  const Workbook* workbook() const noexcept { return workbook_; }

  /// Returns the active lexical-scope environment for bare name references
  /// (`LET` bindings; eventually LAMBDA parameters). Null indicates the
  /// top-level scope where no LET is in flight: under that shape every
  /// `NameRef` resolves to `#NAME?` because defined-name lookup at workbook
  /// scope is not yet wired.
  const NameEnv* name_env() const noexcept { return name_env_; }

  /// Returns a copy of `*this` whose `name_env()` is `env`. Used by the LET
  /// evaluator to extend scope for each binding initialiser and the body
  /// without touching the parent context (which may be shared between
  /// sibling evaluations).
  EvalContext with_name_env(const NameEnv* env) const noexcept {
    EvalContext copy = *this;
    copy.name_env_ = env;
    return copy;
  }

  /// Sentinel row / column value meaning "no formula cell is bound". Chosen
  /// beyond `Sheet::kMaxRows` / `kMaxCols` so it never collides with a valid
  /// 0-based address.
  static constexpr std::uint32_t kNoFormulaCell = static_cast<std::uint32_t>(-1);

  /// Returns a copy of `*this` anchored at the formula cell located at
  /// (row, col) 0-based on the current sheet. Zero-argument position
  /// functions (ROW(), COLUMN()) read this address to return the row / column
  /// of the cell that owns the formula. During recursive `resolve_ref`
  /// evaluation the anchor is updated to the resolved target cell so that
  /// ROW() / COLUMN() inside a referenced formula report the referenced
  /// cell's coordinates, not the caller's.
  EvalContext with_formula_cell(std::uint32_t row, std::uint32_t col) const noexcept {
    EvalContext copy = *this;
    copy.formula_row_ = row;
    copy.formula_col_ = col;
    return copy;
  }

  /// Returns the 0-based row of the formula cell that owns the currently
  /// evaluated expression, or `kNoFormulaCell` when no cell is bound (e.g.
  /// ad-hoc expression evaluation in the CLI).
  std::uint32_t formula_row() const noexcept { return formula_row_; }

  /// Returns the 0-based column of the formula cell that owns the currently
  /// evaluated expression, or `kNoFormulaCell` when no cell is bound.
  std::uint32_t formula_col() const noexcept { return formula_col_; }

  /// True when the context is anchored at a specific formula cell.
  bool has_formula_cell() const noexcept { return formula_row_ != kNoFormulaCell; }

 private:
  const Sheet* current_sheet_ = nullptr;
  EvalState* state_ = nullptr;
  const Workbook* workbook_ = nullptr;
  const NameEnv* name_env_ = nullptr;
  std::uint32_t formula_row_ = kNoFormulaCell;
  std::uint32_t formula_col_ = kNoFormulaCell;
};

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_EVAL_CONTEXT_H_
