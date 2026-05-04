// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the parallel SCC-layered recalc scheduler. See
// `scheduler.h` for the public contract; this TU owns the layering
// algorithm, the worker-pool dispatch, and the workbook-level write
// mutex.

#include "eval/scheduler.h"

#include <algorithm>
#include <atomic>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
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
#include "eval/volatile_tracker.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Hard cap on auto-detected worker count. Matches the `PTHREAD_POOL_SIZE`
// the WASM build reserves so behaviour is consistent across native and
// browser runtimes.
constexpr std::uint32_t kMaxAutoThreads = 8U;

// Per-thread re-entrancy flag. Set to true on entry to `recalc_parallel`
// and cleared on exit. A nested invocation on the same thread (typically a
// UDF that re-enters the engine via `Workbook::recalc_parallel`) trips the
// flag and surfaces `kGraphRecalcReentrant`. The scheduler intentionally
// does NOT extend the check across worker threads: separate threads are
// still expected to obey the documented "no two concurrent
// `recalc_parallel` on the same workbook" contract; this guard exists to
// fail fast on a recursive same-thread call rather than to police
// cross-thread misuse.
thread_local bool g_in_recalc = false;

// RAII guard that scopes the `g_in_recalc` flag to a recalc invocation.
// Captures the prior value so it restores correctly even when nested
// invocations are blocked further up the stack (the outer call releases
// the flag and the next sibling invocation proceeds normally).
struct ReentrantGuard {
  bool prev;
  ReentrantGuard() noexcept : prev(g_in_recalc) { g_in_recalc = true; }
  ~ReentrantGuard() { g_in_recalc = prev; }
  ReentrantGuard(const ReentrantGuard&) = delete;
  ReentrantGuard& operator=(const ReentrantGuard&) = delete;
  ReentrantGuard(ReentrantGuard&&) = delete;
  ReentrantGuard& operator=(ReentrantGuard&&) = delete;
};

// Returns true when the SCC `component` is cyclic (more than one cell, or
// a singleton with a self-loop). Mirrors `recalc_engine.cpp`'s helper —
// kept private here to avoid a public dependency.
bool is_cyclic_component(const std::vector<CellNodeId>& component, const DepGraph& graph) {
  if (component.size() > 1U) {
    return true;
  }
  const CellNodeId only = component.front();
  std::vector<CellNodeId> deps = graph.dependencies_of(only);
  return std::find(deps.begin(), deps.end(), only) != deps.end();
}

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
      for (CellNodeId dep : graph.dependencies_of(member)) {
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

ThreadArenas make_thread_arenas(std::size_t count) {
  ThreadArenas arenas;
  arenas.reserve(count);
  for (std::size_t i = 0; i < count; ++i) {
    arenas.emplace_back(std::make_unique<Arena>());
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
};

SccOutcome process_scc(const std::vector<CellNodeId>& component, Workbook& wb, const DepGraph& graph,
                       const FunctionRegistry& registry, const IterativeOptions& iter_opts, Arena& arena,
                       std::mutex& write_mutex) {
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
        sheet.set_cell_cached_value(c.row, c.col, Value::error(ErrorCode::Ref));
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

    const IterativeOutcome outcome = run_iterative_solve(component, iter_opts, evaluate_one, commit);
    if (outcome.converged) {
      ++out.cycle_recoveries;
      out.cells_evaluated += component.size();
    } else if (!outcome.diverged) {
      // Iteration-limit exhaustion: write #NUM! ourselves so the cells
      // do not retain misleading partial values.
      for (CellNodeId c : component) {
        if (c.sheet_id >= sheet_count) {
          continue;
        }
        Sheet& sheet = wb.sheet(c.sheet_id);
        std::lock_guard<std::mutex> guard(write_mutex);
        sheet.set_cell_cached_value(c.row, c.col, Value::error(ErrorCode::Num));
      }
    }
    // Divergence path already wrote #NUM! through the solver's commit.
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
  Value result = evaluate_cell_for_recalc(wb, sheet, *cell_data, only.row, only.col, registry, arena);
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
  ThreadArenas* arenas = nullptr;
  std::mutex* write_mutex = nullptr;
  std::vector<SccOutcome>* outcomes = nullptr;
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
    (*work->outcomes)[idx] = process_scc((*work->components)[(*work->tasks)[idx]], *work->wb, *work->graph,
                                         *work->registry, *work->iter_opts, arena, *work->write_mutex);
  }
}

// Result of a single layer pool drive. Distinguishes "all workers
// started, every task processed" from "thread spawn failed mid-pool"
// so the caller can serial-fallback only the unprocessed tail rather
// than blindly re-running every task.
struct LayerPoolResult {
  /// True iff every requested worker thread started successfully.
  bool spawn_ok = false;
  /// Number of tasks the worker pool actually claimed via `next_index`,
  /// clamped to the total task count. Tasks `[claimed, total)` were
  /// never picked up and must be processed by the caller's fallback.
  std::size_t claimed = 0U;
};

// Runs the worker pool for a single layer. All started threads are
// joined before returning — `detach()` is forbidden by project policy.
//
// On a partial spawn failure (a `std::thread` constructor returned a
// non-joinable handle, typically because the OS refused another
// thread), the started workers still drain as much of the queue as
// they can; `next_index` reports how many tasks they collectively
// claimed so the caller can resume from there.
LayerPoolResult run_layer_pool(std::uint32_t num_threads, LayerWork& work) {
  work.outcomes->assign(work.tasks->size(), SccOutcome{});
  std::vector<std::thread> workers;
  workers.reserve(num_threads);
  LayerPoolResult result;
  result.spawn_ok = true;
  for (std::uint32_t i = 0; i < num_threads; ++i) {
    std::thread t(worker_loop, static_cast<std::size_t>(i), &work);
    if (!t.joinable()) {
      result.spawn_ok = false;
      break;
    }
    workers.push_back(std::move(t));
  }
  for (std::thread& t : workers) {
    if (t.joinable()) {
      t.join();
    }
  }
  // Snapshot `next_index` AFTER every worker has joined. The acq_rel
  // updates inside the loop pair with this load, so we observe the
  // last-claimed-plus-one count published by whichever worker drained
  // the queue last. Clamp to the task total because the final claim
  // may have stepped past the end (workers fetch then check).
  const std::size_t raw_claimed = work.next_index.load(std::memory_order_acquire);
  result.claimed = std::min(raw_claimed, work.tasks->size());
  return result;
}

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
  ReentrantGuard guard;

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

  // Phase 3: Tarjan SCC over the entire graph (same as the serial engine).
  const std::vector<std::vector<CellNodeId>> all_sccs = engine.graph_.tarjan_scc();

  // Filter to dirty SCCs and assign condensed-graph indices.
  SccIndex idx;
  idx.components.reserve(all_sccs.size());
  for (const std::vector<CellNodeId>& comp : all_sccs) {
    bool any_dirty = false;
    for (CellNodeId c : comp) {
      if (engine.dirty_.contains(c)) {
        any_dirty = true;
        break;
      }
    }
    if (!any_dirty) {
      continue;
    }
    const std::size_t scc_id = idx.components.size();
    for (CellNodeId c : comp) {
      idx.cell_to_scc.emplace(c, scc_id);
    }
    idx.components.push_back(comp);
  }

  const CondensedGraph cg = build_condensed_graph(idx, engine.graph_);
  const std::vector<std::vector<std::size_t>> layers = kahn_layers(cg);

  const std::uint32_t worker_count = resolve_thread_count(cfg.num_threads);
  ThreadArenas arenas = make_thread_arenas(worker_count);
  std::mutex write_mutex;

  std::uint64_t cells_evaluated = 0;
  std::uint64_t sccs_processed = 0;
  std::uint64_t parallel_steps = 0;
  std::uint64_t serial_fallback_steps = 0;
  std::uint64_t cycle_recoveries = 0;

  const IterativeOptions iter_opts = engine.iterative_options();

  for (const std::vector<std::size_t>& layer : layers) {
    sccs_processed += layer.size();

    if (layer.size() <= 1U || worker_count <= 1U) {
      // Tiny layer or single-worker mode: process serially on this thread.
      ++serial_fallback_steps;
      Arena& arena = *arenas[0U];
      for (std::size_t scc_id : layer) {
        const SccOutcome o =
            process_scc(idx.components[scc_id], wb, engine.graph_, registry, iter_opts, arena, write_mutex);
        cells_evaluated += o.cells_evaluated;
        cycle_recoveries += o.cycle_recoveries;
      }
      continue;
    }

    // Parallel layer: cap workers at the layer size (no point spawning
    // 8 threads to drain a 2-task layer).
    const std::uint32_t threads_for_layer =
        std::min<std::uint32_t>(worker_count, static_cast<std::uint32_t>(layer.size()));
    std::vector<SccOutcome> outcomes;
    LayerWork work;
    work.tasks = &layer;
    work.components = &idx.components;
    work.wb = &wb;
    work.graph = &engine.graph_;
    work.registry = &registry;
    work.iter_opts = &iter_opts;
    work.arenas = &arenas;
    work.write_mutex = &write_mutex;
    work.outcomes = &outcomes;
    const LayerPoolResult pool_result = run_layer_pool(threads_for_layer, work);

    for (const SccOutcome& o : outcomes) {
      cells_evaluated += o.cells_evaluated;
      cycle_recoveries += o.cycle_recoveries;
    }
    if (pool_result.spawn_ok) {
      ++parallel_steps;
    } else {
      // Partial spawn failure: drain the unprocessed tail synchronously
      // on the calling thread. `pool_result.claimed` is the exact count
      // of tasks the started workers picked up via `next_index`, so we
      // process `[claimed, layer.size())` without re-running anything
      // the pool already finished. This relies on workers honouring
      // their fetch_add bookkeeping even when peers failed to spawn,
      // which they do: the failure mode aborts only the pool builder,
      // not the workers already in flight.
      ++serial_fallback_steps;
      Arena& arena = *arenas[0U];
      for (std::size_t i = pool_result.claimed; i < layer.size(); ++i) {
        const SccOutcome o =
            process_scc(idx.components[layer[i]], wb, engine.graph_, registry, iter_opts, arena, write_mutex);
        cells_evaluated += o.cells_evaluated;
        cycle_recoveries += o.cycle_recoveries;
      }
    }
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
      Value result = evaluate_cell_for_recalc(wb, sheet, *cell_data, c.row, c.col, registry, arena);
      sheet.set_cell_cached_value(c.row, c.col, result);
      ++cells_evaluated;
      ++sccs_processed;  // Treat a standalone formula as its own component.
    }
    ++serial_fallback_steps;
  }

  // Phase 5: clear the dirty set.
  engine.dirty_.clear();

  if (stats != nullptr) {
    stats->cells_evaluated = cells_evaluated;
    stats->sccs_processed = sccs_processed;
    stats->parallel_steps = parallel_steps;
    stats->serial_fallback_steps = serial_fallback_steps;
    stats->cycle_recoveries = cycle_recoveries;
  }

  return Expected<void, Error>::Ok();
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
