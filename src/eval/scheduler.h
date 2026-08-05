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
// one another, so they may be evaluated concurrently. A pool of
// `std::thread` workers drains a per-layer queue; each worker holds its
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

#include <cstdint>

#include "utils/error.h"
#include "utils/expected.h"

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
  /// Layers dispatched on the worker pool (i.e. layer size >= 2 AND the
  /// pool was successfully spawned).
  std::uint64_t parallel_steps = 0;
  /// Layers processed serially on the calling thread because the layer
  /// had <= 1 super-node or the caller selected one worker.
  std::uint64_t serial_fallback_steps = 0;
  /// Number of cyclic SCCs successfully resolved by the iterative solver
  /// (excluding `#REF!`-fallback components).
  std::uint64_t cycle_recoveries = 0;
};

/// Drives a parallel incremental recalc on `wb`.
///
/// The function reads `wb`'s embedded `RecalcEngine` to obtain the dep
/// graph + dirty / volatile state, performs the same end-to-end pass as
/// `RecalcEngine::recalc`, and clears the dirty set on completion.
///
/// `cfg` selects the worker count (see `SchedulerConfig::num_threads`).
/// `stats` (when non-null) receives the per-pass counters. The `Expected`
/// return slot is reserved for scheduler failures. A requested worker pool
/// is created with `std::thread`; on no-exception builds an operating-system
/// thread-creation failure cannot be recovered by this API. Callers that
/// require serial execution must set `SchedulerConfig::num_threads` to 1.
Expected<void, Error> recalc_parallel(Workbook& wb, const SchedulerConfig& cfg = {}, SchedulerStats* stats = nullptr);

/// Convenience overload that runs `recalc_parallel` against a caller-
/// provided registry. Default callers should prefer the 1-arg overload
/// which threads in `default_registry()` automatically.
Expected<void, Error> recalc_parallel(Workbook& wb, const FunctionRegistry& registry, const SchedulerConfig& cfg = {},
                                      SchedulerStats* stats = nullptr);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SCHEDULER_H_
