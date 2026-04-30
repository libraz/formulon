// Copyright 2026 libraz. Licensed under the MIT License.
//
// Workbook-wide incremental recalculation engine.
//
// `RecalcEngine` is the per-`Workbook` orchestrator that owns:
//
//   * a `DepGraph` recording every `formula -> cell` edge in the workbook;
//   * a `VolatileTracker` listing formulas whose AST contains an Excel
//     volatile function (NOW, RAND, OFFSET, ...);
//   * a `DirtySet` accumulating cells that need to be re-evaluated on the
//     next pass.
//
// The engine itself is single-threaded; the multi-threaded scheduler that
// will eventually parallelise SCC evaluation lives in a later phase. For
// now the contract is intentionally narrow:
//
//   * `register_formula(cell, ast, workbook)` analyses `ast` (via
//     `extract_deps`) and updates the dep graph + volatile tracker for
//     `cell`. It is idempotent: re-registering a cell drops its previous
//     edges before adding the new ones.
//   * `unregister_formula(cell)` drops every edge owned by the cell and
//     deregisters it from the volatile set. Used when a cell becomes a
//     literal or is cleared.
//   * `mark_dirty(cell)` adds the cell to the `DirtySet`. Callers (the
//     `Workbook` mutation API) invoke it whenever a cell's value or
//     formula text changes.
//   * `recalc(workbook, registry)` performs an end-to-end incremental
//     recalc: it seeds the dirty set with volatile cells, BFS-propagates
//     dirtiness through reverse edges, runs Tarjan SCC over the *full*
//     graph (so the reverse-topological order is preserved), and evaluates
//     every dirty SCC in order. Singleton SCCs without a self-loop are
//     evaluated via the tree walker; SCCs with cycles surface `#REF!` on
//     every member (until the iterative solver lands in a follow-up
//     bundle). The dirty set is cleared at the end of the pass.

#ifndef FORMULON_EVAL_RECALC_ENGINE_H_
#define FORMULON_EVAL_RECALC_ENGINE_H_

#include <cstdint>
#include <memory>

#include "eval/dep_graph.h"
#include "eval/dirty_set.h"
#include "eval/iterative_solver.h"
#include "eval/volatile_tracker.h"
#include "parser/ast.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Arena;
class Workbook;

namespace eval {

class FunctionRegistry;
class RecalcEngine;
struct SchedulerConfig;
struct SchedulerStats;

// Free-function entry point of the parallel scheduler. Declared here so
// the implementation in `scheduler.cpp` can be made a friend of
// `RecalcEngine` and reach its private graph / dirty / volatile state
// without further widening the engine's public surface.
Expected<void, Error> recalc_parallel_impl(Workbook& wb, const FunctionRegistry& registry, const SchedulerConfig& cfg,
                                           SchedulerStats* stats, RecalcEngine& engine);

/// Statistics describing a single `recalc()` pass. Useful for tests and
/// observability; the recalc engine itself does not interpret these counts.
struct RecalcStats {
  /// Total cells whose formula was executed (including volatile-driven
  /// re-evaluations). Cells in cycle SCCs are NOT counted here — they were
  /// short-circuited to `#REF!` without invoking the evaluator.
  std::uint32_t cells_evaluated = 0;
  /// Cells that were forced into the dirty set by their volatility status.
  /// A volatile cell whose value did not change is still counted because
  /// the formula was re-executed.
  std::uint32_t volatile_cells = 0;
  /// Cells that participated in a cycle SCC and could not be resolved.
  /// Two cases feed this counter:
  ///   * iterative calc disabled — every cyclic SCC's members surface
  ///     `#REF!` and bump this count;
  ///   * iterative calc enabled — only SCCs whose iterative solve failed
  ///     (max-iteration exhaustion or divergence detection) bump this
  ///     count, with each member receiving `#NUM!`.
  /// Successful iterative solves count toward `iterative_cells` /
  /// `cells_evaluated` instead.
  std::uint32_t cycle_cells = 0;
  /// Cells whose value was computed by the iterative solver (i.e. they
  /// participated in a cyclic SCC that converged with iterative calc
  /// enabled). Each such cell is also counted in `cells_evaluated` —
  /// `iterative_cells <= cells_evaluated` — so callers that want a
  /// "cells touched by iterative-calc machinery" total can read the
  /// dedicated counter without double-counting plain singleton evals.
  std::uint32_t iterative_cells = 0;
  /// Reserved for cooperative cancellation. Always false until the cancel
  /// token surface lands.
  bool cancelled = false;
};

/// Single-threaded incremental recalc orchestrator. Owned 1:1 by a
/// `Workbook`; not copyable, not movable across workbooks.
class RecalcEngine {
 public:
  RecalcEngine();
  ~RecalcEngine();

  RecalcEngine(const RecalcEngine&) = delete;
  RecalcEngine& operator=(const RecalcEngine&) = delete;
  RecalcEngine(RecalcEngine&&) = delete;
  RecalcEngine& operator=(RecalcEngine&&) = delete;

  /// Re-analyses the formula at `cell` and updates the dep graph / volatile
  /// tracker. Drops every previous outgoing edge of `cell` first, so
  /// repeated calls are idempotent. `workbook` is read-only here — it is
  /// only consulted to resolve sheet qualifiers via
  /// `Workbook::sheet_index_by_name`.
  void register_formula(CellNodeId cell, const parser::AstNode& ast, const Workbook& workbook);

  /// Drops every edge owned by `cell` and removes it from the volatile
  /// tracker. Safe to call on a cell that was never registered (no-op).
  void unregister_formula(CellNodeId cell);

  /// Drops only the outgoing dep-graph edges of `cell` (the cells it
  /// reads) and removes it from the volatile tracker. Incoming edges —
  /// the cells that read `cell` — are preserved so existing dependents
  /// keep their wiring. Used by `Workbook::set_cell_value` when a formula
  /// cell is overwritten with a literal: dependents must continue to
  /// re-evaluate against the new literal value.
  void clear_cell_dependencies(CellNodeId cell);

  /// Marks `cell` as dirty for the next `recalc()` pass. Idempotent.
  void mark_dirty(CellNodeId cell);

  /// Performs a full incremental recalc against `workbook`.
  ///
  /// The pass proceeds in five phases:
  ///   1. Seed the dirty set with every volatile cell so they re-execute
  ///      regardless of upstream changes.
  ///   2. BFS-propagate dirtiness via `DepGraph::dependents_of` until no new
  ///      cells are added.
  ///   3. Compute Tarjan SCCs of the entire graph. Tarjan emits leaves
  ///      first, which is exactly the order the evaluator wants (a cell's
  ///      dependencies are evaluated before the cell itself).
  ///   4. Walk the SCC list and evaluate every component whose intersection
  ///      with the dirty set is non-empty. Singleton SCCs without a
  ///      self-loop run through the tree walker; cyclic SCCs receive
  ///      `#REF!` on every member.
  ///   5. Clear the dirty set.
  ///
  /// Returns `RecalcStats` describing the pass. The engine never propagates
  /// errors out of `recalc()`; bad formulas surface as Excel error sentinels
  /// in the cell store. The `Expected` return slot is reserved for future
  /// cancellation / resource-limit failures.
  Expected<RecalcStats, Error> recalc(Workbook& workbook, const FunctionRegistry& registry);

  // ------------------------------------------------------------------------
  // Test / debug accessors. Const-only so external code cannot mutate the
  // engine's internal state without going through the public API.
  // ------------------------------------------------------------------------

  /// Returns the current dependency graph. Lifetime is tied to the engine.
  const DepGraph& dep_graph() const noexcept { return graph_; }
  /// Returns the current volatile tracker.
  const VolatileTracker& volatiles() const noexcept { return volatiles_; }
  /// Returns the current dirty set (post-recalc this is empty).
  const DirtySet& dirty() const noexcept { return dirty_; }

  /// Replaces the iterative-calc options. Default state matches Excel's
  /// out-of-the-box settings (iterative calc disabled, 100 iterations,
  /// 0.001 max change). Callers flip `enabled = true` to opt circular
  /// SCCs into fixed-point resolution; the next `recalc()` pass picks up
  /// the new options.
  void set_iterative_options(IterativeOptions opts) noexcept { iterative_ = opts; }

  /// Returns the active iterative-calc options.
  const IterativeOptions& iterative_options() const noexcept { return iterative_; }

 private:
  // The parallel scheduler reads `graph_` / `volatiles_` and mutates
  // `dirty_` / `arena_` via the same algorithm as `recalc()`. Granting
  // friendship rather than widening the public surface keeps the
  // single-threaded contract intact for every other caller.
  friend Expected<void, Error> recalc_parallel_impl(Workbook&, const FunctionRegistry&, const SchedulerConfig&,
                                                    SchedulerStats*, RecalcEngine&);

  DepGraph graph_;
  VolatileTracker volatiles_;
  DirtySet dirty_;
  // Reused across `recalc()` calls so the bump-allocator's largest chunk
  // sticks around. `reset()` is invoked at the end of every pass.
  std::unique_ptr<Arena> arena_;
  // Iterative-calc knobs. Default-disabled so existing callers keep the
  // legacy `#REF!` behaviour for cyclic SCCs without opting in.
  IterativeOptions iterative_;
};

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_RECALC_ENGINE_H_
