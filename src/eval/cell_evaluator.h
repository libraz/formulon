// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Shared "parse + evaluate one formula cell" helper used by the recalc
// engine and the parallel scheduler.
//
// The cell-evaluation logic — strip leading '=', parse the formula text
// into an AST under `arena`, build an `EvalContext` (workbook-bound, with
// per-call `EvalState` unless iterative mode flips it off), drive
// `evaluate()`, then commit any top-level Array spill back through
// `EvalContext::dispatch_array_result` — was previously duplicated bit-
// for-bit between `recalc_engine.cpp` and `scheduler.cpp`. Both TUs
// now route through `evaluate_cell_for_recalc` so the two recalc entry
// points cannot drift.
//
// This helper does NOT acquire `RecalcEngine::mutex_` and does NOT touch
// the cell store: callers are responsible for synchronising writes back
// to the sheet (the scheduler holds a per-workbook write mutex; the
// serial engine holds the engine mutex for the whole pass).

#ifndef FORMULON_EVAL_CELL_EVALUATOR_H_
#define FORMULON_EVAL_CELL_EVALUATOR_H_

#include <cstdint>

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
  /// When `true`, the EvalContext is built WITHOUT an `EvalState` binding
  /// (`EvalContext::workbook_only`). Formula-cell reads then short-circuit
  /// to `Cell::cached_value` instead of recursing into the cell's formula
  /// text — exactly what the iterative solver wants so each member of an
  /// SCC reads its peers' most-recently-committed values rather than
  /// triggering re-entrant evaluation.
  bool iterative_mode = false;
};

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
