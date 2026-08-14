//
// Implementation of the parallel SCC-layered recalc scheduler. See
// `scheduler.h` for the public contract; this TU owns the layering
// algorithm, the worker-pool dispatch, and the workbook-level write
// mutex.

#include "eval/scheduler.h"

#include <algorithm>
#include <atomic>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <string>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/cell_evaluator.h"
#include "eval/dep_graph.h"
#include "eval/dirty_set.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "eval/recalc_reentry.h"
#include "eval/spill_release.h"
#include "eval/volatile_tracker.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "utils/structured_log.h"
#include "utils/thread_launch.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Hard cap on auto-detected worker count. Matches the `PTHREAD_POOL_SIZE`
// the WASM build reserves so behaviour is consistent across native and
// browser runtimes.
constexpr std::uint32_t kMaxAutoThreads = 8U;

// Per-thread re-entrancy flag and guard live in `recalc_reentry.h` so the
// serial `RecalcEngine::recalc` and the parallel scheduler share the same
// flag. A nested invocation on the same thread (typically a UDF or a
// progress callback that calls back into the engine) trips the flag and
// surfaces `kGraphRecalcReentrant` instead of deadlocking on the engine
// mutex. The scheduler intentionally does NOT extend the check across
// worker threads: separate threads are still expected to obey the
// documented "no two concurrent recalc on the same workbook" contract;
// this guard exists to fail fast on a recursive same-thread call rather
// than to police cross-thread misuse.
using detail::g_in_recalc;
using detail::RecalcReentryGuard;

// The spill-release queue, the progress snapshot and the wave ceilings are
// shared with the serial `RecalcEngine::recalc`; see `spill_release.h`.
using detail::BlockedSpillState;
using detail::canonical_release_targets;
using detail::kMaxNoProgressSpillWaves;
using detail::kMaxSpillReleaseWaves;
using detail::queue_spill_release;
using detail::snapshot_blocked_spills;
using detail::SpillReleaseQueue;

// Maps each cell in the dirty SCC list to its 0-based super-node id. SCC
// indices are assigned in iteration order of `sccs_dirty`.
struct SccIndex {
  std::vector<std::vector<CellNodeId>> components;  // dense list of dirty SCCs
  std::unordered_map<CellNodeId, std::size_t, CellNodeIdHash> cell_to_scc;
};

// Builds the inter-SCC condensed graph adjacency restricted to the dirty
// component set. `predecessors[i]` holds the dirty-SCC indices whose
// members are read by some member of SCC `i`; `successors[i]` is the
// transpose. Self-loops (a cell inside SCC `i` reading another member of
// the same SCC) are NOT recorded — those are interior to the cycle and
// handled by the iterative solver.
struct CondensedGraph {
  std::vector<std::vector<std::size_t>> predecessors;
  std::vector<std::vector<std::size_t>> successors;
  std::vector<std::size_t> in_degree;
};

CondensedGraph build_condensed_graph(const SccIndex& idx, const DepGraph& graph) {
  CondensedGraph out;
  const std::size_t n = idx.components.size();
  out.predecessors.resize(n);
  out.successors.resize(n);
  out.in_degree.assign(n, 0U);

  // Per-SCC dedup buffer; reused across iterations.
  std::unordered_set<std::size_t> seen_preds;

  for (std::size_t i = 0; i < n; ++i) {
    seen_preds.clear();
    for (CellNodeId member : idx.components[i]) {
      for (CellNodeId dep : graph.dependencies_of_ref(member)) {
        auto it = idx.cell_to_scc.find(dep);
        if (it == idx.cell_to_scc.end()) {
          continue;  // Dependency outside the dirty set: irrelevant for layering.
        }
        const std::size_t j = it->second;
        if (j == i) {
          continue;  // Internal SCC edge — solver concern, not scheduler concern.
        }
        if (seen_preds.insert(j).second) {
          out.predecessors[i].push_back(j);
          out.successors[j].push_back(i);
          ++out.in_degree[i];
        }
      }
    }
  }

  return out;
}

// Computes layers via Kahn's algorithm. `layers[k]` holds the SCC indices
// whose dependencies all live in layers `< k`. Because the input graph is
// the dirty-SCC condensation (which is a DAG by construction), this
// always succeeds: every node ends up in some layer.
std::vector<std::vector<std::size_t>> kahn_layers(const CondensedGraph& cg) {
  std::vector<std::vector<std::size_t>> layers;
  std::vector<std::size_t> in_deg = cg.in_degree;
  std::vector<std::size_t> frontier;
  frontier.reserve(in_deg.size());
  for (std::size_t i = 0; i < in_deg.size(); ++i) {
    if (in_deg[i] == 0U) {
      frontier.push_back(i);
    }
  }

  std::vector<std::size_t> next_frontier;
  while (!frontier.empty()) {
    layers.push_back(frontier);
    next_frontier.clear();
    for (std::size_t i : frontier) {
      for (std::size_t j : cg.successors[i]) {
        if (in_deg[j] > 0U) {
          --in_deg[j];
          if (in_deg[j] == 0U) {
            next_frontier.push_back(j);
          }
        }
      }
    }
    frontier.swap(next_frontier);
  }

  return layers;
}

// Per-thread evaluation arenas. Allocated up front (one per worker) and
// reset per task — `Arena::reset` keeps the largest chunk so steady-state
// tasks become allocation-free. Stored as `unique_ptr` so the addresses
// stay stable across vector growth (which we never trigger because the
// pool is sized once at scheduler entry, but the cheap insurance is
// worth the indirection).
using ThreadArenas = std::vector<std::unique_ptr<Arena>>;

ThreadArenas make_thread_arenas(std::size_t count, std::size_t max_arena_bytes) {
  ThreadArenas arenas;
  arenas.reserve(count);
  for (std::size_t i = 0; i < count; ++i) {
    arenas.emplace_back(std::make_unique<Arena>(/*initial_chunk_bytes=*/4096, max_arena_bytes));
  }
  return arenas;
}

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

SccOutcome process_scc(const std::vector<CellNodeId>& component, Workbook& wb, const DepGraph& graph,
                       const FunctionRegistry& registry, const IterativeOptions& iter_opts, Arena& arena,
                       IterativeProgressCb progress_cb, void* progress_user_data, std::mutex& write_mutex,
                       SpillReleaseCallback release_callback, void* release_user_data) {
  SccOutcome out;
  const std::size_t sheet_count = wb.sheet_count();

  if (is_cyclic_component(component, graph)) {
    if (!iter_opts.enabled) {
      // Cycle SCC, iterative calc disabled: every member surfaces #REF!.
      for (CellNodeId c : component) {
        if (c.sheet_id >= sheet_count) {
          continue;
        }
        Sheet& sheet = wb.sheet(c.sheet_id);
        std::lock_guard<std::mutex> guard(write_mutex);
        SpillCommitter committer(&sheet, c.row, c.col, release_callback, release_user_data);
        sheet.set_cell_cached_value(c.row, c.col, committer.commit(Value::error(ErrorCode::Ref)));
      }
      return out;
    }

    // Iterative calc enabled: drive the solver. The lambdas below capture
    // by reference and run synchronously inside this worker — the solver
    // is single-threaded and performs the full SCC fixed-point search
    // before returning. Concurrent SCC processors operate on disjoint
    // cells (different SCCs by construction), so the cell-store mutex
    // only ever contends inside `commit`.
    auto evaluate_one = [&](CellNodeId c) -> Value {
      if (c.sheet_id >= sheet_count) {
        return Value::error(ErrorCode::Ref);
      }
      Sheet& sheet = wb.sheet(c.sheet_id);
      const Cell* cell_data = sheet.cell_at(c.row, c.col);
      if (cell_data == nullptr || cell_data->formula_text.empty()) {
        return Value::blank();
      }
      arena.reset();
      EvaluateCellOptions opts;
      opts.iterative_mode = true;
      opts.spill_release_callback = release_callback;
      opts.spill_release_user_data = release_user_data;
      return evaluate_cell_for_recalc(wb, sheet, *cell_data, c.row, c.col, registry, arena, opts);
    };
    auto commit = [&](CellNodeId c, Value v) {
      if (c.sheet_id >= sheet_count) {
        return;
      }
      Sheet& sheet = wb.sheet(c.sheet_id);
      std::lock_guard<std::mutex> guard(write_mutex);
      sheet.set_cell_cached_value(c.row, c.col, v);
    };

    const IterativeOutcome outcome =
        run_iterative_solve(component, iter_opts, evaluate_one, commit, progress_cb, progress_user_data);
    if (arena.exhausted()) {
      out.arena_exhausted = true;
      return out;
    }
    if (outcome.converged) {
      ++out.cycle_recoveries;
      out.cells_evaluated += component.size();
    }
    return out;
  }

  // Plain singleton.
  const CellNodeId only = component.front();
  if (only.sheet_id >= sheet_count) {
    return out;
  }
  Sheet& sheet = wb.sheet(only.sheet_id);
  const Cell* cell_data = sheet.cell_at(only.row, only.col);
  if (cell_data == nullptr || cell_data->formula_text.empty()) {
    return out;
  }
  arena.reset();
  EvaluateCellOptions opts;
  opts.spill_release_callback = release_callback;
  opts.spill_release_user_data = release_user_data;
  Value result = evaluate_cell_for_recalc(wb, sheet, *cell_data, only.row, only.col, registry, arena, opts);
  if (arena.exhausted()) {
    out.arena_exhausted = true;
    return out;
  }
  {
    std::lock_guard<std::mutex> guard(write_mutex);
    sheet.set_cell_cached_value(only.row, only.col, result);
  }
  ++out.cells_evaluated;
  return out;
}

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

// Drains the layer's task queue from one worker. `outcomes` is
// preallocated so each worker writes into its own slot without a
// result-side lock. `write_mutex` serialises only the per-cell store
// mutation, not the evaluation itself.
//
// `next_index` is updated with `acq_rel`: the acquire half ensures any
// writes a peer worker performed before claiming its task (in particular,
// the `outcomes[idx] = ...` store) are visible to subsequent observers
// of `next_index`; the release half publishes this worker's own writes
// before the next claim. `relaxed` would have been correct only because
// the join below provides the happens-before edge to the caller, but a
// future change that observes `next_index` outside the join would silently
// race; `acq_rel` keeps the invariant local to this loop.
void worker_loop(std::size_t worker_id, LayerWork* work) {
  Arena& arena = *(*work->arenas)[worker_id];
  while (true) {
    const std::size_t idx = work->next_index.fetch_add(1U, std::memory_order_acq_rel);
    if (idx >= work->tasks->size()) {
      return;
    }
    // Each worker owns one byte in this vector. The scheduler reads the
    // flags only after the layer barrier (and, ultimately, after the pool is
    // joined), so distinct workers can set their own flags without a result
    // lock or an atomic increment on the hot task-claim path.
    if (work->worker_used != nullptr) {
      (*work->worker_used)[worker_id] = 1U;
    }
    (*work->outcomes)[idx] =
        process_scc((*work->components)[(*work->tasks)[idx]], *work->wb, *work->graph, *work->registry,
                    *work->iter_opts, arena, work->progress_cb, work->progress_user_data, *work->write_mutex,
                    work->release_callback, work->release_user_data);
  }
}

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
  explicit LayerWorkerPool(std::uint32_t worker_count) : worker_used_(worker_count, 0U) {
    // Both vectors are reserved and never grown afterwards: each live
    // worker holds a pointer into `slots_`, which a reallocation would
    // dangle.
    slots_.reserve(worker_count);
    workers_.reserve(worker_count);
    for (std::uint32_t worker_id = 0U; worker_id < worker_count; ++worker_id) {
      slots_.push_back(WorkerSlot{this, worker_id, ThreadStart{}});
      WorkerSlot& slot = slots_.back();
      slot.start.entry = &LayerWorkerPool::worker_entry;
      slot.start.arg = &slot;
      auto thread_or = launch_thread(slot.start);
      if (!thread_or) {
        // Degrade rather than abort: drop the slot that never started and
        // run the pass on the workers already alive.
        slots_.pop_back();
        StructuredLog("recalc.worker.launch_failed")
            .field("requested", static_cast<std::int64_t>(worker_count))
            .field("started", static_cast<std::int64_t>(workers_.size()))
            .field("reason", thread_or.error().message)
            .warn();
        break;
      }
      workers_.push_back(std::move(thread_or.value()));
    }
  }

  LayerWorkerPool(const LayerWorkerPool&) = delete;
  LayerWorkerPool& operator=(const LayerWorkerPool&) = delete;

  ~LayerWorkerPool() {
    {
      std::lock_guard<std::mutex> lock(mutex_);
      stopping_ = true;
    }
    work_ready_.notify_all();
    for (Thread& worker : workers_) {
      worker.join();
    }
  }

  /// Workers actually running, which is at most what the constructor was
  /// asked for.
  std::uint32_t size() const noexcept { return static_cast<std::uint32_t>(workers_.size()); }

  /// Number of workers that claimed at least one task. Must be called after
  /// the caller has drained all layers (or after destruction), because worker
  /// flags are intentionally non-atomic and are published by the layer
  /// barrier / thread joins.
  std::uint32_t used_count() const noexcept {
    std::uint32_t count = 0U;
    for (const std::uint8_t used : worker_used_) {
      count += used != 0U ? 1U : 0U;
    }
    return count;
  }

  std::vector<std::uint8_t>* used_flags() noexcept { return &worker_used_; }

  /// Dispatches one layer across the live workers and returns once they
  /// have all come back to the barrier. Requires `size() >= 1`: with no
  /// worker to drain the queue this would return with the layer
  /// untouched, so the caller checks the count first and takes the serial
  /// path instead.
  void run(LayerWork& work) {
    work.outcomes->assign(work.tasks->size(), SccOutcome{});
    work.next_index.store(0U, std::memory_order_relaxed);
    std::unique_lock<std::mutex> lock(mutex_);
    active_work_ = &work;
    completed_workers_ = 0U;
    ++generation_;
    work_ready_.notify_all();
    layer_done_.wait(lock, [this] { return completed_workers_ == workers_.size(); });
    active_work_ = nullptr;
  }

 private:
  /// One launched worker's identity, kept at a stable address for the
  /// lifetime of the thread that reads it.
  struct WorkerSlot {
    LayerWorkerPool* pool = nullptr;
    std::uint32_t worker_id = 0U;
    ThreadStart start;
  };

  static void worker_entry(void* raw) {
    auto* slot = static_cast<WorkerSlot*>(raw);
    // The re-entry guard is thread-local. The caller's guard therefore does
    // not protect callbacks running on a pool worker; install a worker-side
    // guard so a nested serial or parallel recalc fails before trying to
    // acquire the already-held engine mutex.
    RecalcReentryGuard guard;
    slot->pool->worker_main(slot->worker_id);
  }

  void worker_main(std::uint32_t worker_id) {
    std::size_t seen_generation = 0U;
    while (true) {
      LayerWork* work = nullptr;
      {
        std::unique_lock<std::mutex> lock(mutex_);
        work_ready_.wait(lock, [this, seen_generation] { return stopping_ || generation_ != seen_generation; });
        if (stopping_) {
          return;
        }
        seen_generation = generation_;
        work = active_work_;
      }
      worker_loop(static_cast<std::size_t>(worker_id), work);
      {
        std::lock_guard<std::mutex> lock(mutex_);
        ++completed_workers_;
        if (completed_workers_ == workers_.size()) {
          layer_done_.notify_one();
        }
      }
    }
  }

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

std::uint32_t resolve_thread_count(std::uint32_t requested) {
  if (requested == 0U) {
    const unsigned hw = std::thread::hardware_concurrency();
    if (hw == 0U) {
      return 2U;  // Fallback when the runtime cannot probe the CPU.
    }
    return std::min<std::uint32_t>(static_cast<std::uint32_t>(hw), kMaxAutoThreads);
  }
  return std::min<std::uint32_t>(requested, kMaxAutoThreads);
}

}  // namespace

// ---------------------------------------------------------------------------
// recalc_parallel_impl — friend of RecalcEngine, declared in recalc_engine.h.
// ---------------------------------------------------------------------------
Expected<void, Error> recalc_parallel_impl(Workbook& wb, const FunctionRegistry& registry, const SchedulerConfig& cfg,
                                           SchedulerStats* stats, RecalcEngine& engine) {
  // Keep the public scheduler contract bounded as well as the C ABI wrapper:
  // values above the eight-worker ceiling are invalid, not silently rounded
  // down. The output is already reset before any validation or re-entry
  // check, so every failure path leaves a caller-provided stats object clean.
  if (stats != nullptr) {
    *stats = SchedulerStats{};
  }
  if (cfg.num_threads > kMaxAutoThreads) {
    return make_error(FormulonErrorCode::kInvalidArgument, "recalc_parallel thread count exceeds the scheduler limit",
                      "thread_count=" + std::to_string(cfg.num_threads) + " max=8");
  }

  // Re-entrancy check. If this thread is already inside a
  // `recalc_parallel` invocation, surface `kGraphRecalcReentrant` instead
  // of corrupting the dirty / dep-graph bookkeeping. Triggered today only
  // by a UDF that calls back into the engine; the guard exists so a future
  // host-callable function family cannot accidentally race the dirty set
  // by attempting nested recalc.
  if (g_in_recalc) {
    return make_error(FormulonErrorCode::kGraphRecalcReentrant, "recalc_parallel called recursively on the same thread",
                      "the scheduler does not support nested recalc; the inner call is rejected");
  }
  RecalcReentryGuard guard;

  // Acquire the engine mutex for the entire pass. Workers spawned below
  // operate on `Sheet` storage (guarded by `write_mutex` for the cell
  // store) and read-only views of the engine's internal state, so a
  // single critical section over the whole recalc keeps the dep graph /
  // dirty set / volatile tracker consistent against any concurrent
  // `register_formula` / `mark_dirty` issued by another thread. Same-
  // thread re-entry is already short-circuited by `g_in_recalc` above,
  // so a non-recursive `std::mutex` is sufficient — UDFs invoked during
  // evaluation must not call back into mutating `RecalcEngine` APIs.
  std::lock_guard<std::mutex> engine_guard(engine.mutex_);

  SpillReleaseQueue release_queue{&wb};
  const SpillReleaseCallback release_callback = &queue_spill_release;
  std::uint64_t cells_evaluated = 0;
  std::uint64_t sccs_processed = 0;
  std::uint64_t scc_nodes_considered = 0;
  std::uint64_t parallel_steps = 0;
  std::uint64_t serial_fallback_steps = 0;
  std::uint64_t cycle_recoveries = 0;
  std::size_t release_waves = 0U;
  std::size_t dependency_waves = 0U;
  std::size_t no_progress_waves = 0U;
  std::vector<BlockedSpillState> previous_release_state;
  std::vector<CellNodeId> previous_release_targets;
  bool have_previous_release_state = false;
  const std::uint32_t configured_worker_count = resolve_thread_count(cfg.num_threads);
  // Keep one pool for the complete pass. A spill/dependency retry is another
  // wave of the same recalc, not a new recalc invocation; workers therefore
  // remain parked at the layer barrier until the final wave is complete.
  std::unique_ptr<LayerWorkerPool> worker_pool;
  ThreadArenas arenas = make_thread_arenas(1U, cfg.max_arena_bytes);
  std::mutex write_mutex;
  const auto mark_release_targets_dirty = [&](const std::vector<CellNodeId>& anchors) {
    for (const CellNodeId anchor : anchors) {
      engine.dirty_.mark(anchor);
      for (const CellNodeId dependent : engine.graph_.dependents_of(anchor)) {
        engine.dirty_.mark(dependent);
      }
      engine.mark_range_dependents_dirty_locked(anchor);
    }
  };

  for (;;) {
    // Reconcile committed spill geometry before closure/SCC scheduling. A
    // producer rewrite may have removed a previously valid phantom edge;
    // dirtying its watcher before Tarjan avoids a stale cycle. Added and
    // removed source ownership both wake the watcher at this boundary.
    const DepGraph::DependencyDelta pre_dependency_delta = engine.reconcile_spill_dependencies_locked(wb);
    for (const DepGraph::Edge& edge : pre_dependency_delta.added) {
      engine.dirty_.mark(edge.first);
    }
    for (const DepGraph::Edge& edge : pre_dependency_delta.removed) {
      engine.dirty_.mark(edge.first);
    }

    // Phase 1 + 2: seed the dirty set with volatile cells, BFS-propagate
    // dirtiness through reverse edges. Mirrors `RecalcEngine::recalc`.
    engine.volatiles_.for_each([&](CellNodeId cell) {
      if (!engine.dirty_.contains(cell)) {
        engine.dirty_.mark(cell);
      }
    });

    std::vector<CellNodeId> bfs_queue;
    bfs_queue.reserve(engine.dirty_.size());
    engine.dirty_.for_each([&](CellNodeId c) { bfs_queue.push_back(c); });
    std::size_t bfs_head = 0;
    while (bfs_head < bfs_queue.size()) {
      const CellNodeId current = bfs_queue[bfs_head++];
      for (CellNodeId dependent : engine.graph_.dependents_of(current)) {
        if (!engine.dirty_.contains(dependent)) {
          engine.dirty_.mark(dependent);
          bfs_queue.push_back(dependent);
        }
      }
    }

    // Phase 3: Tarjan SCC over the dirty induced subgraph, exactly the node
    // set the serial engine passes to `tarjan_scc_subset`. Running it over
    // the whole registered graph instead would make every wave of the retry
    // loop below cost the workbook rather than the edit.
    std::unordered_set<CellNodeId, CellNodeIdHash> dirty_nodes;
    dirty_nodes.reserve(engine.dirty_.size());
    engine.dirty_.for_each([&](CellNodeId c) { dirty_nodes.insert(c); });
    const std::vector<std::vector<CellNodeId>> dirty_sccs = engine.graph_.tarjan_scc_subset(dirty_nodes);
    scc_nodes_considered += dirty_nodes.size();

    // Assign condensed-graph indices. Every emitted component is dirty by
    // construction, so no filtering pass is needed.
    SccIndex idx;
    idx.components.reserve(dirty_sccs.size());
    for (const std::vector<CellNodeId>& comp : dirty_sccs) {
      const std::size_t scc_id = idx.components.size();
      for (CellNodeId c : comp) {
        idx.cell_to_scc.emplace(c, scc_id);
      }
      idx.components.push_back(comp);
    }

    const CondensedGraph cg = build_condensed_graph(idx, engine.graph_);
    const std::vector<std::vector<std::size_t>> layers = kahn_layers(cg);

    const IterativeOptions iter_opts = engine.iterative_options();
    const IterativeProgressCb progress_cb = engine.progress_cb_;
    void* const progress_user_data = engine.progress_user_data_;

    for (const std::vector<std::size_t>& layer : layers) {
      sccs_processed += layer.size();

      // A progress callback is user code and must stay on the caller. It is
      // enough for one cyclic SCC to occur in the layer to force the whole
      // barrier through the caller, preserving serial callback affinity and
      // avoiding callback races between otherwise independent SCCs.
      bool callback_cycle = false;
      if (progress_cb != nullptr) {
        for (const std::size_t scc_id : layer) {
          if (is_cyclic_component(idx.components[scc_id], engine.graph_)) {
            callback_cycle = true;
            break;
          }
        }
      }

      bool run_serial = layer.size() <= 1U || configured_worker_count <= 1U || callback_cycle;
      if (!run_serial && worker_pool == nullptr) {
        // The pool is lazy so a no-op / serial-only recalc starts no OS
        // workers merely because the caller selected a parallel cap. Once
        // created, this object is retained across all subsequent waves.
        worker_pool = std::make_unique<LayerWorkerPool>(configured_worker_count);
        arenas = make_thread_arenas(std::max<std::uint32_t>(worker_pool->size(), 1U), cfg.max_arena_bytes);
      }
      if (!run_serial && worker_pool->size() <= 1U) {
        // A launch refusal is a successful degradation, not a scheduler
        // failure. A single surviving worker cannot make a layer parallel,
        // so keep the caller-only path. Keep the pool (even when empty) so a
        // later spill wave does not retry the OS launch and accidentally
        // create a second pool in this pass.
        run_serial = true;
      }

      if (run_serial) {
        // Tiny layer or single-worker mode: process serially on this thread.
        ++serial_fallback_steps;
        Arena& arena = *arenas[0U];
        for (std::size_t scc_id : layer) {
          const SccOutcome o =
              process_scc(idx.components[scc_id], wb, engine.graph_, registry, iter_opts, arena, progress_cb,
                          progress_user_data, write_mutex, release_callback, &release_queue);
          if (o.arena_exhausted) {
            return make_error(FormulonErrorCode::kOutOfMemory, "evaluation arena exhausted during parallel recalc");
          }
          cells_evaluated += o.cells_evaluated;
          cycle_recoveries += o.cycle_recoveries;
        }
        continue;
      }

      // Parallel layer: the workers are already alive and return to the
      // barrier for the next layer after draining this queue.
      std::vector<SccOutcome> outcomes;
      LayerWork work;
      work.tasks = &layer;
      work.components = &idx.components;
      work.wb = &wb;
      work.graph = &engine.graph_;
      work.registry = &registry;
      work.iter_opts = &iter_opts;
      work.progress_cb = progress_cb;
      work.progress_user_data = progress_user_data;
      work.arenas = &arenas;
      work.write_mutex = &write_mutex;
      work.release_callback = release_callback;
      work.release_user_data = &release_queue;
      work.outcomes = &outcomes;
      work.worker_used = worker_pool->used_flags();
      worker_pool->run(work);

      for (const SccOutcome& o : outcomes) {
        if (o.arena_exhausted) {
          return make_error(FormulonErrorCode::kOutOfMemory, "evaluation arena exhausted during parallel recalc");
        }
        cells_evaluated += o.cells_evaluated;
        cycle_recoveries += o.cycle_recoveries;
      }
      ++parallel_steps;
    }

    // Phase 4b: standalone-dirty cells (no graph edges, e.g. `=NOW()`).
    // These are not in any SCC; sweep them serially. The fast `cells_to_scc`
    // map already covers every cell that landed in a dirty SCC.
    std::vector<CellNodeId> standalone_dirty;
    engine.dirty_.for_each([&](CellNodeId c) {
      if (idx.cell_to_scc.find(c) == idx.cell_to_scc.end()) {
        standalone_dirty.push_back(c);
      }
    });
    if (!standalone_dirty.empty()) {
      Arena& arena = *arenas[0U];
      const std::size_t sheet_count = wb.sheet_count();
      for (CellNodeId c : standalone_dirty) {
        if (c.sheet_id >= sheet_count) {
          continue;
        }
        Sheet& sheet = wb.sheet(c.sheet_id);
        const Cell* cell_data = sheet.cell_at(c.row, c.col);
        if (cell_data == nullptr || cell_data->formula_text.empty()) {
          continue;
        }
        arena.reset();
        EvaluateCellOptions opts;
        opts.spill_release_callback = release_callback;
        opts.spill_release_user_data = &release_queue;
        Value result = evaluate_cell_for_recalc(wb, sheet, *cell_data, c.row, c.col, registry, arena, opts);
        if (arena.exhausted()) {
          return make_error(FormulonErrorCode::kOutOfMemory, "evaluation arena exhausted during parallel recalc");
        }
        sheet.set_cell_cached_value(c.row, c.col, result);
        ++cells_evaluated;
        ++sccs_processed;  // Treat a standalone formula as its own component.
      }
      ++serial_fallback_steps;
    }

    const DepGraph::DependencyDelta dependency_delta = engine.reconcile_spill_dependencies_locked(wb);
    const bool dependency_retry = !dependency_delta.added.empty();
    if (dependency_retry) {
      ++dependency_waves;
      for (const DepGraph::Edge& edge : dependency_delta.added) {
        engine.dirty_.mark(edge.first);
      }
    }

    const std::vector<CellNodeId> released = release_queue.take();
    if (!released.empty() || dependency_retry) {
      ++release_waves;
      if (!released.empty()) {
        const std::vector<BlockedSpillState> release_state = snapshot_blocked_spills(wb);
        const std::vector<CellNodeId> release_targets = canonical_release_targets(released);
        if (have_previous_release_state && release_state == previous_release_state &&
            release_targets == previous_release_targets) {
          ++no_progress_waves;
        } else {
          no_progress_waves = 0U;
        }
        previous_release_state = release_state;
        previous_release_targets = release_targets;
        have_previous_release_state = true;
      }
      if (release_waves > kMaxSpillReleaseWaves || dependency_waves > kMaxSpillReleaseWaves ||
          (!released.empty() && no_progress_waves >= kMaxNoProgressSpillWaves)) {
        // Keep the current dirty set, including unrelated work, while
        // retaining the release targets for a caller retry after an
        // external mutation.
        mark_release_targets_dirty(released);
        return make_error(FormulonErrorCode::kGraphScheduleFailed, "spill release waves made no progress",
                          "parallel dynamic-array spill recovery exceeded its bounded wave budget");
      }
      engine.dirty_.clear();
      mark_release_targets_dirty(released);
      for (const DepGraph::Edge& edge : dependency_delta.added) {
        engine.dirty_.mark(edge.first);
      }
      continue;
    }

    // Phase 5: clear the dirty set.
    engine.dirty_.clear();

    if (stats != nullptr) {
      stats->cells_evaluated = cells_evaluated;
      stats->sccs_processed = sccs_processed;
      stats->scc_nodes_considered = scc_nodes_considered;
      stats->parallel_steps = parallel_steps;
      stats->serial_fallback_steps = serial_fallback_steps;
      stats->cycle_recoveries = cycle_recoveries;
      if (worker_pool != nullptr) {
        stats->worker_threads_started = worker_pool->size();
        stats->worker_threads_used = worker_pool->used_count();
      }
    }

    return Expected<void, Error>::Ok();
  }
}

Expected<void, Error> recalc_parallel(Workbook& wb, const SchedulerConfig& cfg, SchedulerStats* stats) {
  return recalc_parallel(wb, default_registry(), cfg, stats);
}

Expected<void, Error> recalc_parallel(Workbook& wb, const FunctionRegistry& registry, const SchedulerConfig& cfg,
                                      SchedulerStats* stats) {
  // The engine is owned by the workbook; the scheduler is a friend of
  // RecalcEngine so it may reach the private dep-graph / dirty / volatile
  // state through `recalc_parallel_impl`.
  return recalc_parallel_impl(wb, registry, cfg, stats, wb.recalc_engine());
}

}  // namespace formulon::eval
