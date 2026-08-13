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
#include <mutex>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include "eval/dep_extractor.h"
#include "eval/dep_graph.h"
#include "eval/dirty_set.h"
#include "eval/iterative_solver.h"
#include "eval/range_dep_index.h"
#include "eval/spill_potential.h"
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

/// Workbook-relative half-open viewport rectangle, expressed in 0-based
/// inclusive coordinates. Used by `RecalcEngine::partial_recalc` to bound
/// the work performed during a latency-sensitive UI redraw.
///
/// `last_row` / `last_col` are inclusive — a single-cell viewport sets
/// `first_*` and `last_*` to the same value. An empty viewport (the row
/// range or the column range collapsed) is allowed and produces a
/// no-op recalc; the engine treats it as "the UI does not need any
/// fresh values right now".
struct SheetCellRange {
  /// 0-based sheet index. Out-of-range values are silently ignored by
  /// `partial_recalc`.
  std::uint16_t sheet_id = 0;
  /// First row, 0-based, inclusive.
  std::uint32_t first_row = 0;
  /// Last row, 0-based, inclusive.
  std::uint32_t last_row = 0;
  /// First column, 0-based, inclusive.
  std::uint32_t first_col = 0;
  /// Last column, 0-based, inclusive.
  std::uint32_t last_col = 0;
};

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
};

/// Single-threaded incremental recalc orchestrator. Owned 1:1 by a
/// `Workbook`; not copyable, not movable across workbooks.
///
/// Threading model
/// ---------------
/// Mutating public APIs (`register_formula`, `unregister_formula`,
/// `clear_cell_dependencies`, `mark_dirty`, every `recalc*` entry) acquire
/// the internal `mutex_` for the full duration of the call. Internal
/// `*_locked` helpers assume the lock is already held by the caller; the
/// parallel scheduler reaches them through `friend` access after taking
/// `mutex_` itself, so workers within a single recalc pass run while the
/// engine state is held under one continuous critical section.
///
/// During a recalc pass user-supplied callbacks (UDFs, the iterative-solver
/// progress callback) MUST NOT invoke any mutating `RecalcEngine` API or
/// any `Workbook` method that would route into one — doing so would deadlock
/// (`mutex_` is a non-recursive `std::mutex`). Same-thread re-entry into
/// `recalc_parallel` is additionally rejected up-front by the scheduler's
/// `thread_local` re-entrancy guard.
///
/// Workbook-wide lock order: workbook compound mutation -> RecalcEngine
/// mutex -> Sheet spill mutex (and, for parallel evaluation, the scheduler's
/// write mutex before the Sheet spill mutex). Spill snapshots release the
/// Sheet lock before graph replacement; the graph is immutable while workers
/// run.
class RecalcEngine {
 public:
  RecalcEngine();
  ~RecalcEngine();

  RecalcEngine(const RecalcEngine&) = delete;
  RecalcEngine& operator=(const RecalcEngine&) = delete;
  RecalcEngine(RecalcEngine&&) = delete;
  RecalcEngine& operator=(RecalcEngine&&) = delete;

  /// Passkey-style facade that exposes the engine's `*_locked` mutators
  /// to `Workbook`'s compound-mutation entry points. The constructor is
  /// friend-restricted to `RecalcEngine` / `Workbook`, so a caller can
  /// only obtain a `LockedMutator` from inside a `Workbook` member
  /// function that has already taken `engine.mutex_`. The facade itself
  /// does NOT acquire or release the lock — callers MUST hold a
  /// `std::lock_guard<std::mutex>` on `RecalcEngine::mutex_for_compound_mutation()`
  /// for the full duration of every call routed through this object.
  ///
  /// The type is publicly named so anonymous-namespace helpers inside
  /// `workbook.cpp` can take it by reference; only `Workbook` can mint
  /// one, so the surface area stays controlled.
  class LockedMutator {
   public:
    LockedMutator(const LockedMutator&) = delete;
    LockedMutator& operator=(const LockedMutator&) = delete;
    LockedMutator(LockedMutator&&) = delete;
    LockedMutator& operator=(LockedMutator&&) = delete;
    ~LockedMutator() = default;

    void register_formula(CellNodeId cell, const parser::AstNode& ast, const Workbook& workbook) const;
    void unregister_formula(CellNodeId cell) const;
    void clear_cell_dependencies(CellNodeId cell) const;
    void mark_dirty(CellNodeId cell) const;
    /// Marks formulas that own a compact whole-row / whole-column
    /// dependency covering `cell` dirty. Called alongside the ordinary
    /// reverse-edge walk for every workbook cell mutation.
    void mark_range_dependents_dirty(CellNodeId cell) const;
    /// Drops the entire dependency graph, volatile set, and dirty set.
    /// Used by `Workbook`'s sheet-permutation entry points, which
    /// invalidate every `CellNodeId.sheet_id` at once and must re-register
    /// every formula from scratch rather than patch individual edges.
    void reset_graph() const;
    /// Returns the formula owners whose extracted Ref3D span includes
    /// `edited_sheet`. The caller receives a stable snapshot while the
    /// workbook mutation lock is held; callers may then remap owners after
    /// the physical row/column move.
    std::vector<CellNodeId> three_d_span_owners_covering_sheet(std::uint16_t edited_sheet) const;
    const DepGraph& dep_graph() const noexcept;

   private:
    friend class RecalcEngine;
    friend class ::formulon::Workbook;
    explicit LockedMutator(RecalcEngine& engine) noexcept : engine_(engine) {}

    RecalcEngine& engine_;
  };

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

  /// Recalculates only those dirty cells whose final result is required to
  /// produce correct values within the supplied viewport rectangle.
  ///
  /// Computes the dependency closure backward from the viewport: every cell
  /// transitively reachable from a viewport cell via the dependency graph is
  /// included; cells outside that closure are skipped even if they are
  /// otherwise dirty. Volatile cells inside the closure are still
  /// recomputed every call. Cells outside the closure remain dirty; a
  /// subsequent full `recalc()` (or a `partial_recalc` whose closure
  /// overlaps them) will visit them.
  ///
  /// Edge cases:
  ///   * Empty viewport (`first_row > last_row` or `first_col >
  ///     last_col`): no-op, returns `RecalcStats{}` with all zero counts.
  ///   * Viewport extending beyond a sheet's bounds: silently clamped to
  ///     the populated cells inside the rectangle.
  ///   * Out-of-range `sheet_id`: no-op, returns `RecalcStats{}`.
  ///   * Cycles inside the closure: surfaced exactly the way `recalc()`
  ///     surfaces them (`#REF!` when iterative calc is disabled, the
  ///     iterative solver otherwise) — the closure restriction never
  ///     hides a circular reference.
  ///
  /// Threading: this method runs on the calling thread, holding the same
  /// invariants as `recalc()`. Future revisions may dispatch the closure
  /// through the parallel scheduler; today the contract is single-threaded.
  Expected<RecalcStats, Error> partial_recalc(Workbook& workbook, const FunctionRegistry& registry,
                                              const SheetCellRange& viewport);

  // ------------------------------------------------------------------------
  // Test / debug accessors. Const-only so external code cannot mutate the
  // engine's internal state without going through the public API.
  // ------------------------------------------------------------------------

  /// Returns the current dependency graph. Lifetime is tied to the engine.
  const DepGraph& dep_graph() const noexcept { return graph_; }

  /// Returns the cells `cell` reads through a compact rectangle dependency.
  ///
  /// Rectangles above `kMaxMaterializedDependencyCells` never become graph
  /// edges, so a formula-audit consumer that walks `dep_graph()` alone would
  /// not see them at all. The expansion is clipped by
  /// `Sheet::populated_extent` and reports only coordinates that hold content
  /// (stored value / formula, or a committed spill), so a whole-column
  /// reference on a sparse sheet costs the sheet's real content rather than
  /// its 1,048,576 potential cells. Returns an empty vector when `cell` owns
  /// no compact rectangle.
  std::vector<CellNodeId> compact_range_precedents_of(CellNodeId cell, const Workbook& workbook) const;

  /// Returns the formulas that read `cell` through a compact rectangle
  /// dependency — the reverse direction of `compact_range_precedents_of`,
  /// and equally invisible in `dep_graph()`. Order is unspecified;
  /// duplicates are removed.
  std::vector<CellNodeId> compact_range_dependents_of(CellNodeId cell) const;
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

  /// Installs a progress callback for the iterative solver.
  ///
  /// The callback is invoked once per Gauss-Seidel sweep with the
  /// 1-based iteration index, the maximum residual observed during the
  /// sweep, and the configured iteration cap. Returning `false` aborts
  /// the in-flight solve; the recalc accounting then treats the SCC the
  /// same way it treats an iteration-limit exhaustion (members tallied
  /// in `cycle_cells`). Pass `nullptr` to clear the callback.
  ///
  /// `user_data` is forwarded verbatim to every invocation. The engine
  /// does not take ownership; the caller must keep it alive across
  /// every recalc that may consult the callback.
  ///
  /// **Re-entrancy**: the callback MUST NOT call back into any
  /// `RecalcEngine` (`recalc`, `partial_recalc`) or `Workbook` recalc
  /// API on the same thread. The engine guards against accidental
  /// re-entry by surfacing `kGraphRecalcReentrant` from the inner call
  /// (no deadlock, no cell-store corruption), but the caller's own state
  /// is the safe place to mutate from a callback. Cell mutation
  /// (`set_cell_formula`, etc.) issued from the callback is also
  /// unsupported and races with the engine's in-flight bookkeeping.
  void set_iterative_progress(IterativeProgressCb cb, void* user_data) noexcept {
    progress_cb_ = cb;
    progress_user_data_ = user_data;
  }

  /// Returns the currently installed iterative progress callback, or
  /// `nullptr` when none is active.
  IterativeProgressCb iterative_progress() const noexcept { return progress_cb_; }

 private:
  // The parallel scheduler reads `graph_` / `volatiles_` and mutates
  // `dirty_` / `arena_` via the same algorithm as `recalc()`. Granting
  // friendship rather than widening the public surface keeps the
  // single-threaded contract intact for every other caller.
  friend Expected<void, Error> recalc_parallel_impl(Workbook&, const FunctionRegistry&, const SchedulerConfig&,
                                                    SchedulerStats*, RecalcEngine&);

  // `Workbook` mutators (`set_cell_value`, `set_cell_formula`, the
  // row/col edit helpers, `remove_sheet`) issue several engine operations
  // back-to-back and must hold `mutex_` for the entire compound sequence
  // so a concurrent `recalc_parallel` does not observe a half-applied
  // mutation. The friendship lets those mutators take `mutex_` once and
  // call the `*_locked` helpers directly through `LockedMutator`,
  // instead of round-tripping through the public API and re-acquiring
  // the lock between every call.
  friend class ::formulon::Workbook;

  // `Workbook` reaches the locked-mutator facade and the engine mutex
  // through these accessors; they are intentionally private so the rest
  // of the codebase keeps going through the public `recalc()` /
  // `register_*` surface. The friend grant above lets `Workbook`
  // mutators call them when they need a single critical section across
  // multiple engine operations plus a `Sheet` write.
  LockedMutator locked_mutator() noexcept { return LockedMutator(*this); }
  std::mutex& mutex_for_compound_mutation() noexcept { return mutex_; }

  // Internal helpers that assume `mutex_` is already held. Public
  // counterparts above lock and delegate; the scheduler reaches them
  // through friend access after taking `mutex_` itself.
  void register_formula_locked(CellNodeId cell, const parser::AstNode& ast, const Workbook& workbook);
  void unregister_formula_locked(CellNodeId cell);
  void clear_cell_dependencies_locked(CellNodeId cell);
  void mark_dirty_locked(CellNodeId cell);
  void mark_range_dependents_dirty_locked(CellNodeId cell);
  void reset_graph_locked();
  void update_potential_spill_producer_locked(CellNodeId cell, SpillPotential potential);
  Expected<RecalcStats, Error> recalc_locked(Workbook& workbook, const FunctionRegistry& registry);
  Expected<RecalcStats, Error> partial_recalc_locked(Workbook& workbook, const FunctionRegistry& registry,
                                                     const SheetCellRange& viewport);

  /// Rebuilds the spill-footprint-derived portion of the graph from one
  /// snapshot of every sheet's committed spill rectangles. The caller holds
  /// `mutex_`; each sheet lock is acquired only while copying geometry, and
  /// graph mutation happens after those snapshots have been released.
  DepGraph::DependencyDelta reconcile_spill_dependencies_locked(const Workbook& workbook);

  // Serialises every mutating access to `graph_`, `volatiles_`, `dirty_`,
  // and `arena_`. Held for the full duration of each public entry; the
  // parallel scheduler also acquires it once at recalc entry.
  mutable std::mutex mutex_;

  DepGraph graph_;
  // Compact rectangle dependencies (whole-row / whole-column references and
  // every bounded rectangle above `kMaxMaterializedDependencyCells`). Unlike
  // `graph_`, this stores one interned rectangle per authored range rather
  // than one edge per cell.
  RangeDepIndex range_dependencies_;
  struct RegisteredThreeDSpan {
    CellNodeId owner;
    ThreeDSheetSpanDependency span;
  };
  // Ref3D metadata is kept separately from cell/range edges so structural
  // edits can wake every owner whose shared coordinates read the edited
  // sheet, including spans reached through defined-name expansion.
  std::vector<RegisteredThreeDSpan> three_d_span_dependencies_;
  // Formula anchors grouped by sheet whose AST may produce an array. This is
  // the bounded candidate source for partial range expansion; unlike the
  // Sheet row map it never scans unrelated stored cells per wave.
  std::vector<std::unordered_map<CellNodeId, SpillPotential, CellNodeIdHash>> potential_spill_producers_by_sheet_;
  VolatileTracker volatiles_;
  DirtySet dirty_;
  // Reused across `recalc()` calls so the bump-allocator's largest chunk
  // sticks around. `reset()` is invoked at the end of every pass.
  std::unique_ptr<Arena> arena_;
  // Iterative-calc knobs. Default-disabled so existing callers keep the
  // legacy `#REF!` behaviour for cyclic SCCs without opting in.
  IterativeOptions iterative_;
  // Optional progress callback for the iterative solver. `nullptr`
  // disables it (the legacy contract). `progress_user_data_` is
  // forwarded verbatim to every invocation; the engine does not own it.
  IterativeProgressCb progress_cb_ = nullptr;
  void* progress_user_data_ = nullptr;
};

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_RECALC_ENGINE_H_
