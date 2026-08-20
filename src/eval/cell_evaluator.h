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

#include "eval/spill_committer.h"
#include "value.h"

namespace formulon {

class Arena;
struct Cell;
class Sheet;
class Workbook;

namespace eval {

class FunctionRegistry;

/// Knobs for `evaluate_cell_for_recalc`. The defaults reproduce the plain
/// singleton evaluation path used by `RecalcEngine::recalc()`.
struct EvaluateCellOptions {
  SpillReleaseCallback spill_release_callback = nullptr;
  void* spill_release_user_data = nullptr;
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

/// Re-parses and evaluates the formula at `(row, col)` on `sheet`. Returns
/// the result `Value` (or an Excel error sentinel on parse failure). The
/// caller is responsible for writing the result back to the cell.
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
Value evaluate_cell_for_recalc(Workbook& workbook, Sheet& sheet, const Cell& cell_data, std::uint32_t row,
                               std::uint32_t col, const FunctionRegistry& registry, Arena& arena,
                               const EvaluateCellOptions& opts = {});

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_CELL_EVALUATOR_H_
