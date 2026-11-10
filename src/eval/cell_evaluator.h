//
// Shared "parse + evaluate one formula cell" helper used by the recalc
// engine and the parallel scheduler.
//
// The cell-evaluation logic — strip leading '=', parse the formula text
// into an AST under `arena`, build a workbook-bound, state-less
// `EvalContext`, drive `evaluate()`, then commit any top-level Array spill
// back through `EvalContext::dispatch_array_result` — was previously
// duplicated bit-for-bit between `recalc_engine.cpp` and `scheduler.cpp`.
// Both TUs now route through `evaluate_cell_for_recalc` so the two recalc
// entry points cannot drift.
//
// This helper does NOT acquire `RecalcEngine::mutex_` and does NOT touch
// the cell store: callers are responsible for synchronising writes back
// to the sheet (the scheduler holds a per-workbook write mutex; the
// serial engine holds the engine mutex for the whole pass).

#ifndef FORMULON_EVAL_CELL_EVALUATOR_H_
#define FORMULON_EVAL_CELL_EVALUATOR_H_

#include <cstdint>
#include <functional>
#include <mutex>
#include <optional>
#include <vector>

#include "cell.h"
#include "eval/dep_graph.h"
#include "eval/eval_context.h"
#include "eval/spill_committer.h"
#include "sheet_metadata.h"
#include "value.h"

namespace formulon {

class Arena;
class Sheet;
class Workbook;

namespace eval {

class DynamicReadLog;
class FunctionRegistry;

/// Knobs for `evaluate_cell_for_recalc`. The defaults reproduce the plain
/// singleton evaluation path used by `RecalcEngine::recalc()`.
struct EvaluateCellOptions {
  SpillReleaseCallback spill_release_callback = nullptr;
  void* spill_release_user_data = nullptr;
  /// Observer of the rectangles OFFSET / INDIRECT resolve to.
  DynamicReadCallback dynamic_read_callback = nullptr;
  void* dynamic_read_user_data = nullptr;
  /// When set, a state observing reads (`EvalState::observe_reads_of`)
  /// records which of its cells the formula reads.
  EvalState* read_observer = nullptr;
};

/// Copies the formula cell at `(row, col)` out of `sheet` into `staged`,
/// ready to hand to `evaluate_cell_for_recalc`. Returns false when the
/// coordinate holds no formula, in which case `staged` is left untouched.
///
/// Every recalc entry point stages rather than passing `Sheet::cell_at`'s
/// raw `Cell*` straight through, for two independent reasons:
///
///   * Concurrently, a worker holding a `Cell*` across an evaluation would
///     read `formula_text` while a peer's `set_cell_cached_value` rewrites
///     the same slot.
///   * Single-threaded, `cell_at`'s contract forbids retaining the pointer
///     across a mutation that reallocates the row's run, and an evaluation
///     can be one: a refused spill footprint is recorded on the sheet from
///     inside the tree walker. `formula_text` is a bare `std::string`, so a
///     short formula's small-string bytes relocate with the `Cell` itself.
///
/// The staged copy owns the only field the evaluator reads, so neither
/// reaches it.
///
/// `staged` must outlive the `evaluate_cell_for_recalc` call it feeds, and
/// the caller's use of that call's result: the evaluator parses the AST
/// directly out of `staged.formula_text` without interning it, so a string
/// literal in the formula surfaces in the returned `Value` as a view into
/// this buffer. A caller looping over cells therefore declares `staged`
/// outside the loop rather than per iteration.
bool stage_formula_cell(const Sheet& sheet, std::uint32_t row, std::uint32_t col, Cell& staged);

/// One formula cell's evaluation as the recalc driver commits it: the value,
/// and the cell's rich error detail, or `clear` to drop an earlier one.
struct CellResult {
  Value value = Value::blank();
  std::optional<CellErrorDetail> detail;
  bool clear = true;
};

/// Writes `result` to `(row, col)`: the cached value and, with it, the
/// error detail. Every recalc commit of an evaluated formula goes through here.
void commit_cell_result(Sheet& sheet, std::uint32_t row, std::uint32_t col, const CellResult& result);

/// Re-parses and evaluates the formula at `(row, col)` on `sheet`. Returns
/// the result (an Excel error sentinel on parse failure) with its error
/// detail, classified by where the error came from (`eval/error_origin.h`).
/// The caller is responsible for writing the result back to the cell.
///
/// `arena` MUST be safe to mutate from the calling thread — the scheduler
/// uses one arena per worker so concurrent invocations do not race on the
/// bump pointer. `cell_data` is the formula cell as known to the caller;
/// its `formula_text` is expected non-empty (callers check first to avoid
/// the parse round-trip on pure literals).
///
/// If the top-level evaluator produces a `Value::Array` (e.g. a spill
/// anchor returned a `SEQUENCE(...)` result), the spill is committed via
/// `EvalContext::dispatch_array_result` and the anchor scalar is
/// returned, mirroring the behaviour of recursive Array results inside
/// `EvalContext::resolve_ref`.
CellResult evaluate_cell_for_recalc(Workbook& workbook, Sheet& sheet, const Cell& cell_data, std::uint32_t row,
                                    std::uint32_t col, const FunctionRegistry& registry, Arena& arena,
                                    const EvaluateCellOptions& opts = {});

/// Evaluation and commit glue for the members of a cyclic component, shared
/// by the serial engine and the parallel scheduler. `evaluate` stages the
/// member's formula, resets `arena` and runs `evaluate_cell_for_recalc`;
/// `commit` writes a result back to the cell store under `commit_mutex` when
/// one is supplied. A non-null `log` attributes OFFSET / INDIRECT reads to the
/// component at `ordinal`.
class CyclicComponentEvaluator {
 public:
  CyclicComponentEvaluator(Workbook& workbook, const FunctionRegistry& registry, Arena& arena,
                           SpillReleaseCallback release_callback, void* release_user_data, DynamicReadLog* log,
                           std::uint64_t ordinal, std::mutex* commit_mutex)
      : workbook_(workbook),
        registry_(registry),
        arena_(arena),
        release_callback_(release_callback),
        release_user_data_(release_user_data),
        log_(log),
        ordinal_(ordinal),
        commit_mutex_(commit_mutex) {}

  const Workbook& workbook() const noexcept { return workbook_; }

  /// Evaluates `cell` as the recalc driver does, passing `observer` as
  /// `EvaluateCellOptions::read_observer`.
  Value evaluate(CellNodeId cell, EvalState* observer);

  /// Writes `value` to `cell`'s cached value, with the error detail of the
  /// evaluation that produced it when that was the last one, else none.
  void commit(CellNodeId cell, Value value);

 private:
  Workbook& workbook_;
  const FunctionRegistry& registry_;
  Arena& arena_;
  SpillReleaseCallback release_callback_;
  void* release_user_data_;
  DynamicReadLog* log_;
  std::uint64_t ordinal_;
  std::mutex* commit_mutex_;
  // Outlives each evaluation: a string literal in the formula surfaces in the
  // returned Value as a view into this buffer, and the caller consumes each
  // result before asking for the next one (which resets the arena).
  Cell staged_;
  // The last evaluation's cell and detail; drivers commit right after evaluating.
  std::optional<CellNodeId> last_cell_;
  std::optional<CellErrorDetail> last_detail_;
};

/// Settles the cells of a statically cyclic component by what they read (see
/// `settle_cycle_by_reads`), the one rule every recalc driver applies with
/// iterative calculation off. Returns false on a genuine cycle.
bool settle_component_by_reads(const std::vector<CellNodeId>& cells, CyclicComponentEvaluator& evaluator);

/// The value every member cell of a genuine cycle holds with iterative
/// calculation off: 0 for a cell whose cycle runs through itself alone, which
/// Excel leaves at 0 unevaluated (`=A:A` in A1, `=A1+1` in A1), and `#REF!`
/// for a cycle through several cells, whose Excel values depend on the order
/// the cells were entered.
Value unresolved_cycle_value(const std::vector<CellNodeId>& component);

/// Prepares the cached values of an iterative component for its first solve.
/// A cycle persisted while iterative calculation was disabled holds `#REF!`,
/// which is not a useful numerical seed. Until the first full recalculation
/// after iteration is enabled (`RecalcEngine::disabled_cycle_refs_pending`), a
/// formula member with a cached `#REF!` is reset to Blank; afterwards a
/// `#REF!` is a genuine solver result and seeds the solve like any value.
/// Only the supplied component is visited. Callers must hold the recalc
/// write lock while restoring any prior values and invoking this helper.
void prepare_iterative_component_seeds(Workbook& workbook, const std::vector<CellNodeId>& cells);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_CELL_EVALUATOR_H_
