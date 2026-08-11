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

#include "eval/compat.h"
#include "eval/spill_committer.h"
#include "parser/reference.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

class Arena;
class Sheet;
class Workbook;

namespace eval {

class EvalState;
class FunctionRegistry;
class NameEnv;
struct DefinedNameFrame;

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
  EvalContext(const Sheet& sheet, EvalState& state) noexcept : current_sheet_(&sheet), state_(&state) {}

  /// Evaluating + cross-sheet–aware context. Unqualified refs resolve
  /// against `current_sheet`; qualified refs (`Reference::sheet` non-empty)
  /// are looked up in `workbook` case-insensitively. Cycle detection and
  /// memoisation operate on `(sheet, row, col)` via `state`, so cross-sheet
  /// cycles are caught.
  ///
  /// `current_sheet` MUST be a sheet owned by `workbook` (no check is
  /// enforced). `workbook`, `current_sheet`, and `state` must outlive the
  /// context.
  EvalContext(const Workbook& workbook, const Sheet& current_sheet, EvalState& state) noexcept
      : current_sheet_(&current_sheet), state_(&state), workbook_(&workbook) {}

  /// Workbook-aware, *state-less* factory. Unqualified refs resolve
  /// against `current_sheet`, qualified refs are looked up in `workbook`,
  /// and formula cells return their `cached_value` verbatim (the 3-arg
  /// `resolve_ref` overload short-circuits when `state() == nullptr`).
  ///
  /// This is the entry point used by the iterative-calc solver: cyclic
  /// SCC members evaluate against the previous iteration's cached values
  /// rather than triggering re-entrant evaluation that would surface
  /// `#REF!` on any back-edge inside the cycle.
  ///
  /// Both `workbook` and `current_sheet` must outlive the context.
  static EvalContext workbook_only(const Workbook& workbook, const Sheet& current_sheet) noexcept {
    EvalContext out;
    out.current_sheet_ = &current_sheet;
    out.workbook_ = &workbook;
    return out;
  }

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
  Value resolve_ref(const parser::Reference& ref, Arena& arena, const FunctionRegistry& registry) const;

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
  /// Whole-column / whole-row endpoints (`A:A`, `A:C`, `1:1`, `1:3`) are
  /// expanded against the target sheet's used range: the unbounded axis is
  /// clamped to the sheet's populated extent, and the bounded axis keeps
  /// its natural origin (row 0 for a column, column 0 for a row) so
  /// positional consumers (INDEX / VLOOKUP offsets) see the reference's
  /// true top-left. A sheet with no content in range yields an empty
  /// vector. Only same-axis whole references compose; a mixed
  /// whole-column / whole-row pair degrades to `#VALUE!`.
  ///
  /// When non-null, `out_rows` / `out_cols` receive the concrete expanded
  /// shape (both `0` for an empty whole-reference expansion). They are the
  /// only reliable source of shape for whole-reference inputs, whose
  /// clamped dimensions cannot be recovered from the endpoint `Reference`s.
  ///
  /// Short-circuit error mapping (first match wins):
  ///
  /// | Condition                                                 | Result   |
  /// |-----------------------------------------------------------|----------|
  /// | Context is unbound (`current_sheet() == nullptr`)         | `#NAME?` |
  /// | Mismatched cross-sheet endpoints                          | `#REF!`  |
  /// | Qualified ref with no workbook bound                      | `#REF!`  |
  /// | Qualified ref whose target sheet is missing               | `#REF!`  |
  /// | Mixed whole-column / whole-row endpoints                  | `#VALUE!`|
  /// | Either endpoint has row/col >= `Sheet::kMax*`             | `#REF!`  |
  /// | Otherwise                                                 | vector   |
  Expected<std::vector<Value>, ErrorCode> expand_range(const parser::Reference& lhs, const parser::Reference& rhs,
                                                       Arena& arena, const FunctionRegistry& registry,
                                                       std::uint32_t* out_rows = nullptr,
                                                       std::uint32_t* out_cols = nullptr) const;

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

  /// Returns the active Excel formula compatibility profile.
  ExcelProfile excel_profile() const noexcept { return excel_profile_; }

  /// Returns a copy of `*this` with a different Excel formula profile.
  EvalContext with_excel_profile(ExcelProfile profile) const noexcept {
    EvalContext copy = *this;
    copy.excel_profile_ = profile;
    return copy;
  }

  /// True when the bound workbook uses the 1904 date system. Date-aware
  /// evaluators that decompose or compose serials (see
  /// `eval::date_time::serial_from_ymd` / `ymd_from_serial`) must thread
  /// this through so 1904-system workbooks do not shift every date by the
  /// 1462-day epoch gap. Defaults to false (1900 system) for the unbound
  /// and sheet-only context shapes. Sourced from `Workbook::date1904()`
  /// via `with_date1904` at the cell-evaluator boundary.
  bool date1904() const noexcept { return date1904_; }

  /// Returns a copy of `*this` carrying the 1904-date-system flag.
  EvalContext with_date1904(bool value) const noexcept {
    EvalContext copy = *this;
    copy.date1904_ = value;
    return copy;
  }

  /// Returns a copy of `*this` whose recursive-evaluation `state()` is
  /// `&state`. Used by the iterative-calc driver to run each fixed-point
  /// pass against a fresh memoisation map / in-progress stack so stale
  /// memos from the previous pass never leak forward.
  EvalContext with_state(EvalState& state) const noexcept {
    EvalContext copy = *this;
    copy.state_ = &state;
    return copy;
  }

  /// Returns a copy of `*this` with the `evaluate()`-level iterative-calc
  /// fixed-point driver suppressed. The recalc engine owns iterative-calc
  /// resolution itself (via `RecalcEngine` + the SCC iterative solver), so
  /// each per-cell `evaluate()` it issues must run a single pass — not
  /// re-drive a nested fixed-point loop. The direct-`evaluate()` path used
  /// by the CLI and the oracle harness leaves this unset so the
  /// `evaluate()`-level driver runs.
  EvalContext with_iterative_driver_suppressed() const noexcept {
    EvalContext copy = *this;
    copy.suppress_iterative_driver_ = true;
    return copy;
  }

  /// True when the `evaluate()`-level iterative-calc driver is suppressed
  /// for this context (the recalc engine sets this; see
  /// `with_iterative_driver_suppressed`).
  bool iterative_driver_suppressed() const noexcept { return suppress_iterative_driver_; }

  /// Returns a copy of `*this` whose `name_env()` is `env`. Used by the LET
  /// evaluator to extend scope for each binding initialiser and the body
  /// without touching the parent context (which may be shared between
  /// sibling evaluations).
  EvalContext with_name_env(const NameEnv* env) const noexcept {
    EvalContext copy = *this;
    copy.name_env_ = env;
    return copy;
  }

  /// Returns the head of the defined-name expansion chain, or `nullptr` when
  /// no defined name is currently being resolved. The chain is walked by
  /// `resolve_defined_name` (see `defined_name_resolve.h`) to break circular
  /// definitions (`Loop = Loop + 1`) without overflowing the native stack.
  const DefinedNameFrame* defined_name_stack() const noexcept { return defined_name_stack_; }

  /// Returns a copy of `*this` whose `defined_name_stack()` head is `frame`.
  /// `resolve_defined_name` layers one frame per active name expansion; the
  /// frame lives on the resolver's call stack and must outlive every
  /// evaluator call that observes the returned context.
  EvalContext with_defined_name_frame(const DefinedNameFrame* frame) const noexcept {
    EvalContext copy = *this;
    copy.defined_name_stack_ = frame;
    return copy;
  }

  class Builder;

  /// Returns a fluent builder pre-bound to the workbook-aware,
  /// state-carrying shape `(workbook, current_sheet, state)`. Successive
  /// calls (`with_*`) layer on the optional bindings — name environment,
  /// formula-cell anchor, mutable sheet — without forcing each call site
  /// to remember which constructor overload pairs with which combination
  /// of bindings.
  ///
  /// The builder is the recommended construction surface for new call
  /// sites; the legacy public constructors are retained for backward
  /// compatibility and may be migrated incrementally.
  static Builder builder(const Workbook& workbook, const Sheet& current_sheet, EvalState& state) noexcept;

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

  /// Returns a copy of `*this` whose `mutable_sheet()` is `&sheet`. Opting
  /// in to a mutable sheet authorises `dispatch_array_result` to commit
  /// dynamic-array spills on `sheet` for any formula cell anchored on the
  /// same sheet. The base `current_sheet_` (read-only view) is unchanged;
  /// callers that need both a read view and write authority typically pass
  /// the same sheet to both bindings.
  EvalContext with_mutable_sheet(Sheet& sheet) const noexcept {
    EvalContext copy = *this;
    copy.mutable_sheet_ = &sheet;
    return copy;
  }

  /// Installs the recalc-owned observer used to queue blocked anchors when a
  /// committed spill rectangle is released during evaluation.
  EvalContext with_spill_release_callback(SpillReleaseCallback callback, void* user_data) const noexcept {
    EvalContext copy = *this;
    copy.spill_release_callback_ = callback;
    copy.spill_release_user_data_ = user_data;
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

  /// Returns the sheet this context may mutate (for spill commits), or
  /// `nullptr` when no mutable sheet was bound. Distinct from
  /// `current_sheet()`: the latter is the read-only resolution target;
  /// `mutable_sheet()` is the explicit opt-in that authorises spill writes.
  Sheet* mutable_sheet() const noexcept { return mutable_sheet_; }

  /// Returns a pointer to the caller-owned counter that `eval_node` bumps on
  /// entry / decrements on exit. The top-level `evaluate()` entry point
  /// allocates the counter on its own stack frame and points the cloned
  /// context at it; nested `with_*` rebuilders preserve the pointer through
  /// member-wise copy. Null indicates that depth tracking is disabled (for
  /// ad-hoc CLI eval, parser-only smoke tests, and any caller that has not
  /// opted in by going through `evaluate()`).
  std::uint32_t* eval_depth_counter() const noexcept { return eval_depth_counter_; }

  /// Returns a pointer to the caller-owned counter that `invoke_lambda`
  /// bumps on every lambda activation. Same lifetime / null semantics as
  /// `eval_depth_counter()`.
  std::uint32_t* lambda_depth_counter() const noexcept { return lambda_depth_counter_; }

  /// Returns a copy of `*this` with the depth counters bound to the
  /// supplied storage. Both pointers may be null to opt back out of depth
  /// tracking (defensive symmetry; not used in production paths). The
  /// storage must outlive every evaluator call that observes the returned
  /// context.
  EvalContext with_depth_counters(std::uint32_t* eval_counter, std::uint32_t* lambda_counter) const noexcept {
    EvalContext copy = *this;
    copy.eval_depth_counter_ = eval_counter;
    copy.lambda_depth_counter_ = lambda_counter;
    return copy;
  }

  /// Commits an Array result as a dynamic-array spill anchored at the
  /// currently bound formula cell, returning the post-dispatch scalar value
  /// the caller should propagate.
  ///
  /// Behaviour:
  ///   1. Non-Array `v` is returned unchanged (the common scalar path).
  ///   2. If `mutable_sheet()` is null, the caller did not opt in to spill;
  ///      `v` is returned unchanged.
  ///   3. If `has_formula_cell()` is false, there is no anchor address to
  ///      spill into; `v` is returned unchanged.
  ///   4. A degenerate `0 x N` / `N x 0` array (which the producers should
  ///      never emit) yields `#VALUE!` defensively.
  ///   5. Otherwise the array's row-major cells are deep-copied (text
  ///      payloads will be re-interned by `Sheet::commit_spill`) and
  ///      committed at `(formula_row(), formula_col())` on `mutable_sheet()`.
  ///      The return value is `mutable_sheet()->resolve_cell_value(...)` at
  ///      the anchor, which is either `cells[0]` on success or `#SPILL!` on
  ///      collision per the `commit_spill` contract.
  Value dispatch_array_result(Value v) const;

 private:
  const Sheet* current_sheet_ = nullptr;
  EvalState* state_ = nullptr;
  const Workbook* workbook_ = nullptr;
  const NameEnv* name_env_ = nullptr;
  // Head of the intrusive defined-name expansion chain (see
  // `defined_name_resolve.h`). Null at top level; each active name resolution
  // links one frame on so circular definitions are detected instead of
  // recursing without bound.
  const DefinedNameFrame* defined_name_stack_ = nullptr;
  ExcelProfile excel_profile_ = default_excel_profile();
  // 1904 date-system flag, sourced from `Workbook::date1904()`. Threaded
  // to date-aware evaluators so serial <-> calendar conversions pick the
  // correct epoch. Defaults to the 1900 system.
  bool date1904_ = false;
  // Spill-write authority for the current `evaluate()` call. Decoupled from
  // `current_sheet_` so that ad-hoc / read-only contexts (CLI eval, tests
  // that only resolve refs) cannot accidentally mutate the sheet.
  Sheet* mutable_sheet_ = nullptr;
  SpillReleaseCallback spill_release_callback_ = nullptr;
  void* spill_release_user_data_ = nullptr;
  std::uint32_t formula_row_ = kNoFormulaCell;
  std::uint32_t formula_col_ = kNoFormulaCell;
  // Non-owning pointers to caller-stack-allocated depth counters. The
  // top-level `evaluate()` entry point materialises both and binds them
  // via `with_depth_counters`; member-wise copy in the other `with_*`
  // builders preserves the binding for downstream sub-evaluations. Null
  // means "depth tracking disabled" (legacy callers that bypass
  // `evaluate()` retain the pre-existing behaviour).
  std::uint32_t* eval_depth_counter_ = nullptr;
  std::uint32_t* lambda_depth_counter_ = nullptr;
  // When true, the top-level `evaluate()` iterative-calc fixed-point driver
  // is skipped. Set by the recalc engine, which drives iterative calc
  // itself; left false for direct-`evaluate()` callers.
  bool suppress_iterative_driver_ = false;
};

/// Fluent builder for the workbook-aware, state-carrying flavour of
/// `EvalContext`. Construct via `EvalContext::builder(...)` and chain
/// `with_*` calls before terminating with `build()`.
///
/// Rationale: `EvalContext` exposes four public constructors plus the
/// `workbook_only` factory (five distinct shapes). Most call sites need
/// the most-bound shape — workbook + current sheet + state + a name
/// env + a formula-cell anchor + a mutable sheet — and the existing
/// surface forces them to remember the chained `with_*` calls separate
/// from the constructor overload. The builder unifies that into one
/// linear, self-documenting expression. Existing constructor call sites
/// continue to compile unchanged; migration is incremental.
class EvalContext::Builder {
 public:
  Builder(const Workbook& workbook, const Sheet& current_sheet, EvalState& state) noexcept
      : ctx_(workbook, current_sheet, state) {}

  /// Binds the active lexical-scope environment for bare name references
  /// (LET / LAMBDA). Mirrors `EvalContext::with_name_env`.
  Builder& with_name_env(const NameEnv* env) noexcept {
    ctx_ = ctx_.with_name_env(env);
    return *this;
  }

  /// Anchors the context at the formula cell located at `(row, col)` on
  /// the current sheet. Mirrors `EvalContext::with_formula_cell`.
  Builder& with_formula_cell(std::uint32_t row, std::uint32_t col) noexcept {
    ctx_ = ctx_.with_formula_cell(row, col);
    return *this;
  }

  /// Authorises spill commits on `sheet`. Mirrors
  /// `EvalContext::with_mutable_sheet`.
  Builder& with_mutable_sheet(Sheet& sheet) noexcept {
    ctx_ = ctx_.with_mutable_sheet(sheet);
    return *this;
  }

  /// Selects the full Excel formula compatibility profile. Mirrors
  /// `EvalContext::with_excel_profile`.
  Builder& with_excel_profile(ExcelProfile profile) noexcept {
    ctx_ = ctx_.with_excel_profile(profile);
    return *this;
  }

  /// Returns the finished context. The builder may be discarded after
  /// this; further `with_*` calls would have no effect on the returned
  /// value.
  EvalContext build() const noexcept { return ctx_; }

 private:
  EvalContext ctx_;
};

inline EvalContext::Builder EvalContext::builder(const Workbook& workbook, const Sheet& current_sheet,
                                                 EvalState& state) noexcept {
  return Builder(workbook, current_sheet, state);
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_EVAL_CONTEXT_H_
