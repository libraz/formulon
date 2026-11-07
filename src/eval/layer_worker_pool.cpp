//
// Layer worker pool and task-claim loop of the parallel recalc scheduler.

#include "eval/layer_worker_pool.h"

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <mutex>
#include <utility>
#include <vector>

#include "eval/recalc_reentry.h"
#include "utils/structured_log.h"
#include "utils/thread_launch.h"

namespace formulon::eval::detail {
namespace {

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
                    work->release_callback, work->release_user_data, ReadObservation{});
  }
}

}  // namespace

LayerWorkerPool::LayerWorkerPool(std::uint32_t worker_count) : worker_used_(worker_count, 0U) {
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

LayerWorkerPool::~LayerWorkerPool() {
  {
    std::lock_guard<std::mutex> lock(mutex_);
    stopping_ = true;
  }
  work_ready_.notify_all();
  for (Thread& worker : workers_) {
    worker.join();
  }
}

std::uint32_t LayerWorkerPool::used_count() const noexcept {
  std::uint32_t count = 0U;
  for (const std::uint8_t used : worker_used_) {
    count += used != 0U ? 1U : 0U;
  }
  return count;
}

void LayerWorkerPool::run(LayerWork& work) {
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

void LayerWorkerPool::worker_entry(void* raw) {
  auto* slot = static_cast<WorkerSlot*>(raw);
  // The re-entry guard is thread-local. The caller's guard therefore does
  // not protect callbacks running on a pool worker; install a worker-side
  // guard so a nested serial or parallel recalc fails before trying to
  // acquire the already-held engine mutex.
  RecalcReentryGuard guard;
  slot->pool->worker_main(slot->worker_id);
}

void LayerWorkerPool::worker_main(std::uint32_t worker_id) {
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

}  // namespace formulon::eval::detail
