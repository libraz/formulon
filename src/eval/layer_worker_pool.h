//
// Worker pool used by the parallel recalc scheduler, plus the per-layer work
// descriptor and result types it shares with `scheduler.cpp`.

#ifndef FORMULON_EVAL_LAYER_WORKER_POOL_H_
#define FORMULON_EVAL_LAYER_WORKER_POOL_H_

#include <atomic>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <vector>

#include "eval/cell_evaluator.h"
#include "eval/dep_graph.h"
#include "eval/dynamic_read_log.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/spill_release.h"
#include "utils/arena.h"
#include "utils/thread_launch.h"

namespace formulon {
class Workbook;
class Sheet;
namespace eval {
namespace detail {

// Per-thread evaluation arenas. Allocated up front (one per worker) and
// reset per task — `Arena::reset` keeps the largest chunk so steady-state
// tasks become allocation-free. Stored as `unique_ptr` so the addresses
// stay stable across vector growth (which we never trigger because the
// pool is sized once at scheduler entry, but the cheap insurance is
// worth the indirection).
using ThreadArenas = std::vector<std::unique_ptr<Arena>>;

// Evaluates a single dirty SCC and writes the result(s) back into `wb`.
// Holds `write_mutex` only across the per-cell `set_cell_cached_value`
// call so concurrent evaluators on other SCCs of the same layer do not
// serialise on it.
//
// Returns the per-SCC stats contribution. The atomics in
// `recalc_parallel_impl` aggregate them once the layer drains.
struct SccOutcome {
  std::uint64_t cells_evaluated = 0;
  std::uint64_t cycle_recoveries = 0;
  /// Set when the worker's arena reported exhaustion for this component.
  /// The committed values are degraded (an allocation failure surfaces as a
  /// value-level error), so the pass must abort with `kOutOfMemory` rather
  /// than report success — the same contract the serial engine applies at
  /// each of its evaluation sites.
  bool arena_exhausted = false;
};

// How a component's OFFSET / INDIRECT reads are observed. Only components
// run on the calling thread carry a log; pooled ones hold no observed cell
// (see the isolation split in `recalc_parallel_impl`).
struct ReadObservation {
  DynamicReadLog* log = nullptr;
  std::uint64_t ordinal = 0U;

  void observe(CellNodeId cell, EvaluateCellOptions& opts, const Sheet* sheet) const {
    if (log != nullptr) {
      log->observe(cell, ordinal, opts, sheet);
    }
  }
  void note_commit(CellNodeId cell) const {
    if (log != nullptr) {
      log->note_commit(cell, ordinal);
    }
  }
};

// Per-layer shared state passed to each worker by reference. Bundling
// these into a single struct cuts the `std::thread` template
// instantiation surface from ~11 ref arguments down to 2, which trims
// codegen meaningfully.
struct LayerWork {
  std::atomic<std::size_t> next_index{0U};
  const std::vector<std::size_t>* tasks = nullptr;
  const std::vector<std::vector<CellNodeId>>* components = nullptr;
  Workbook* wb = nullptr;
  const DepGraph* graph = nullptr;
  const FunctionRegistry* registry = nullptr;
  const IterativeOptions* iter_opts = nullptr;
  IterativeProgressCb progress_cb = nullptr;
  void* progress_user_data = nullptr;
  ThreadArenas* arenas = nullptr;
  std::mutex* write_mutex = nullptr;
  SpillReleaseCallback release_callback = nullptr;
  void* release_user_data = nullptr;
  std::vector<SccOutcome>* outcomes = nullptr;
  std::vector<std::uint8_t>* worker_used = nullptr;
};

/// Evaluates one dirty component and returns its stats contribution. Defined in
/// `scheduler.cpp`; the worker loop calls it for every claimed task.
SccOutcome process_scc(const std::vector<CellNodeId>& component, Workbook& wb, const DepGraph& graph,
                       const FunctionRegistry& registry, const IterativeOptions& iter_opts, Arena& arena,
                       IterativeProgressCb progress_cb, void* progress_user_data, std::mutex& write_mutex,
                       SpillReleaseCallback release_callback, void* release_user_data, const ReadObservation& obs);

// A worker pool shared by every parallel layer in one recalc pass. Building
// and joining workers per layer dominated small DAGs, so workers wait at a
// layer barrier and are joined once when the pass ends.
//
// The pool takes the requested worker count as an upper bound, not a
// promise. `launch_thread` reports an OS refusal instead of throwing (the
// engine is built `-fno-exceptions`, so `std::thread` would have
// terminated the host here), and the pool keeps whatever workers it did
// start. `size()` is therefore the number to schedule against; the caller
// falls back to evaluating on its own thread when that reaches zero, which
// is slower but produces identical results.
class LayerWorkerPool {
 public:
  explicit LayerWorkerPool(std::uint32_t worker_count);
  LayerWorkerPool(const LayerWorkerPool&) = delete;
  LayerWorkerPool& operator=(const LayerWorkerPool&) = delete;
  ~LayerWorkerPool();

  /// Workers actually running, which is at most what the constructor was
  /// asked for.
  std::uint32_t size() const noexcept { return static_cast<std::uint32_t>(workers_.size()); }
  /// Number of workers that claimed at least one task. Must be called after
  /// the caller has drained all layers (or after destruction), because worker
  /// flags are intentionally non-atomic and are published by the layer
  /// barrier / thread joins.
  std::uint32_t used_count() const noexcept;

  std::vector<std::uint8_t>* used_flags() noexcept { return &worker_used_; }
  /// Dispatches one layer across the live workers and returns once they
  /// have all come back to the barrier. Requires `size() >= 1`: with no
  /// worker to drain the queue this would return with the layer
  /// untouched, so the caller checks the count first and takes the serial
  /// path instead.
  void run(LayerWork& work);

 private:
  /// One launched worker's identity, kept at a stable address for the
  /// lifetime of the thread that reads it.
  struct WorkerSlot {
    LayerWorkerPool* pool = nullptr;
    std::uint32_t worker_id = 0U;
    ThreadStart start;
  };

  static void worker_entry(void* raw);
  void worker_main(std::uint32_t worker_id);

  std::vector<WorkerSlot> slots_;
  std::vector<Thread> workers_;
  std::vector<std::uint8_t> worker_used_;
  std::mutex mutex_;
  std::condition_variable work_ready_;
  std::condition_variable layer_done_;
  LayerWork* active_work_ = nullptr;
  std::size_t generation_ = 0U;
  std::size_t completed_workers_ = 0U;
  bool stopping_ = false;
};

}  // namespace detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LAYER_WORKER_POOL_H_
