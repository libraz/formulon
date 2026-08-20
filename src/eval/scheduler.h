//
// Multi-threaded SCC-parallel recalc scheduler.
//
// `recalc_parallel` is the parallel counterpart of
// `RecalcEngine::recalc`. It performs the same five-phase incremental
// recalc (volatile seed -> dirty BFS -> Tarjan SCC -> evaluate dirty SCCs
// -> clear dirty set), but the "evaluate" phase fans out across a thread
// pool whenever a topological layer contains more than one independent
// super-node.
//
// Every phase is scoped to the same working set as the serial engine: the
// SCC phase runs Tarjan over the dirty induced subgraph, not the registered
// dependency graph, so a small edit to a large workbook costs the edit in
// both engines. Only the evaluation phase differs, and only in how the work
// is distributed.
//
// Algorithm
// ---------
// Once `RecalcEngine::recalc` has computed the SCC list, the scheduler
// rebuilds an inter-SCC condensed graph and partitions it into layers
// using Kahn's algorithm:
//
//   layer 0 -> super-nodes whose dependencies are all outside the dirty
//     working set;
//   layer N -> super-nodes whose dependencies live in layers < N.
//
// Within a single layer no two super-nodes share a dependency edge with
// one another, so they may be evaluated concurrently — with one exception.
// A dynamic reference is the case where the recorded edges do not describe
// the formula's reads: `INDIRECT` / `OFFSET` compute their target at
// evaluation time, so no edge is registered for it and the target may sit
// in the same layer as its reader. A super-node holding such a cell is
// therefore evaluated on the calling thread after the rest of its layer has
// drained, so it reads a cell store no worker is writing. Value-volatile
// cells (`RAND`, `NOW`, ...) keep a complete set of edges and stay on the
// pool; a workbook with no dynamic reference is scheduled exactly as the
// layering describes.
//
// A pool of OS workers drains a per-layer queue; each worker holds its
// own `EvalContext` (the type carries no globals — it is a non-owning
// view of `Workbook` + `Sheet` + `EvalState`). The `Workbook`'s cell
// store is mutated through a single `std::mutex`; the lock is held only
// for the brief window of the per-cell `set_cell_cached_value` call so
// contention is bounded by the per-cell write rate, not the per-cell
// evaluation cost.
//
// Layers with 0 or 1 super-nodes dispatch synchronously on the calling
// thread (skipping the pool overhead). Cyclic SCCs are forwarded to
// `iterative_solver` exactly as in the single-threaded engine; the
// solver runs on whichever worker pulled the SCC off the queue.
//
// Concurrency contract
// --------------------
// Callers MUST NOT race two `recalc_parallel(wb, ...)` calls on the same
// `Workbook` instance. The scheduler does not own a workbook-level lock;
// safety against external readers / writers during a recalc pass remains
// the caller's responsibility (matching `RecalcEngine::recalc`).

#ifndef FORMULON_EVAL_SCHEDULER_H_
#define FORMULON_EVAL_SCHEDULER_H_

#include <cstddef>
#include <cstdint>

#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"

namespace formulon {

class Workbook;

namespace eval {

class FunctionRegistry;

/// Scheduler tuning knobs. Default-constructed values pick automatic
/// thread-count detection (capped at 8) and rely on `default_registry()`
/// for the function dispatch table.
struct SchedulerConfig {
  /// Worker thread count. `0` means auto-detect via
  /// `std::thread::hardware_concurrency()`, clamped to `[1, 8]`. The 8-
  /// thread cap matches the WASM `PTHREAD_POOL_SIZE` ceiling so behaviour
  /// is consistent across native and browser runtimes.
  std::uint32_t num_threads = 0;

  /// Byte ceiling for one worker's evaluation arena. The default matches
  /// the serial engine's single arena, so a cell that evaluates under
  /// `RecalcEngine::recalc` also evaluates here; the pass-wide footprint is
  /// therefore this value times the worker count. Hosts with a tighter
  /// memory budget (wasm32 in particular, whose address space is smaller
  /// than the default ceiling times the worker cap) can lower it: an
  /// evaluation that outgrows the ceiling fails the pass with
  /// `kOutOfMemory` instead of committing a degraded value.
  std::size_t max_arena_bytes = kMaxEvalArenaBytes;
};

/// Counters reported by a `recalc_parallel` invocation. Mirrors the
/// `RecalcStats` shape but adds parallel-specific telemetry (parallel /
/// serial layer counts, cycle recoveries) needed by the scheduler tests.
struct SchedulerStats {
  /// Cells whose formula was actually executed. Mirrors
  /// `RecalcStats::cells_evaluated`.
  std::uint64_t cells_evaluated = 0;
  /// Number of (dirty) SCCs the scheduler processed. Equal to the sum of
  /// per-layer super-node counts walked.
  std::uint64_t sccs_processed = 0;
  /// Nodes submitted to the SCC phase, summed over the pass's waves. This
  /// is the dirty induced subgraph's size, never the registered dependency
  /// graph's, and is reported so that scoping stays observable.
  std::uint64_t scc_nodes_considered = 0;
  /// Layers dispatched on the worker pool (i.e. layer size >= 2 AND the
  /// pool was successfully spawned).
  std::uint64_t parallel_steps = 0;
  /// Serial batches processed on the calling thread: a whole layer that
  /// had <= 1 poolable super-node or ran under a one-worker cap, plus the
  /// dynamic-reference tail of any layer that was otherwise dispatched to
  /// the pool. A layer can therefore contribute to this counter and to
  /// `parallel_steps` at once.
  std::uint64_t serial_fallback_steps = 0;
  /// Number of cyclic SCCs successfully resolved by the iterative solver
  /// (excluding `#REF!`-fallback components).
  std::uint64_t cycle_recoveries = 0;
  /// Number of worker OS threads successfully started for this complete
  /// recalc pass. A caller-only pass reports zero.
  std::uint32_t worker_threads_started = 0;
  /// Number of started workers that claimed at least one SCC task during
  /// this recalc pass. This is a distinct-worker count, not a task count.
  std::uint32_t worker_threads_used = 0;
};

/// Drives a parallel incremental recalc on `wb`.
///
/// The function reads `wb`'s embedded `RecalcEngine` to obtain the dep
/// graph + dirty / volatile state, performs the same end-to-end pass as
/// `RecalcEngine::recalc`, and clears the dirty set on completion.
///
/// `cfg` selects the worker count (see `SchedulerConfig::num_threads`).
/// `stats` (when non-null) receives the per-pass counters.
///
/// A returned `Ok()` means every cell the pass evaluated completed without
/// exhausting its worker's arena. An allocation failure during evaluation
/// aborts the pass with `kOutOfMemory` rather than committing the degraded
/// value, matching `RecalcEngine::recalc`; the remaining `Expected` errors
/// are scheduler failures (invalid worker count, nested recalc, a spill
/// release loop that made no progress).
///
/// `num_threads` is an upper bound, not a guarantee. Workers are started
/// through `launch_thread`, which reports an operating-system refusal
/// rather than throwing, so a host at its thread limit gets a smaller pool
/// instead of a terminated process; a pass that can start no worker at all
/// evaluates every layer on the calling thread and reports each of them in
/// `SchedulerStats::serial_fallback_steps`. The results do not depend on
/// how many workers were obtained. Callers that require serial execution
/// must still set `SchedulerConfig::num_threads` to 1 rather than relying
/// on this degradation.
Expected<void, Error> recalc_parallel(Workbook& wb, const SchedulerConfig& cfg = {}, SchedulerStats* stats = nullptr);

/// Convenience overload that runs `recalc_parallel` against a caller-
/// provided registry. Default callers should prefer the 1-arg overload
/// which threads in `default_registry()` automatically.
Expected<void, Error> recalc_parallel(Workbook& wb, const FunctionRegistry& registry, const SchedulerConfig& cfg = {},
                                      SchedulerStats* stats = nullptr);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SCHEDULER_H_
