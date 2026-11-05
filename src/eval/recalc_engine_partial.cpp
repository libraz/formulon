//
// Viewport-scoped recalculation for `RecalcEngine`: `partial_recalc_locked`.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <unordered_set>
#include <vector>

#include "cell.h"
#include "eval/cell_evaluator.h"
#include "eval/dep_extractor.h"
#include "eval/dep_graph.h"
#include "eval/dynamic_read_log.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "eval/recalc_reentry.h"
#include "eval/spill_release.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/rect_iterator.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

const char* const kArenaExhaustedPartial = "evaluation arena exhausted during partial recalc";

// The spill-release queue, the progress snapshot and the wave ceilings are
// shared with the parallel scheduler; see `spill_release.h`.
using detail::queue_spill_release;
using detail::SpillReleaseQueue;
using detail::SpillWaveBudget;

}  // namespace

Expected<RecalcStats, Error> RecalcEngine::partial_recalc_locked(Workbook& workbook, const FunctionRegistry& registry,
                                                                 const SheetCellRange& viewport) {
  RecalcStats stats;
  SpillReleaseQueue release_queue{&workbook};
  const SpillReleaseCallback release_callback = &queue_spill_release;
  SpillWaveBudget wave_budget;
  DynamicReadPass dynamic(*this, workbook);
  const SerialEvalPass pass{
      workbook, registry, dynamic, release_callback, &release_queue, stats, kArenaExhaustedPartial};

  // ---- Phase 0: validate the viewport. ----
  // Empty viewport — collapsed row / column range, or unknown sheet —
  // is a no-op. The dirty set stays untouched so a subsequent full
  // recalc still picks up everything.
  const std::size_t sheet_count = workbook.sheet_count();
  if (viewport.sheet_id >= sheet_count) {
    return stats;
  }
  if (viewport.first_row > viewport.last_row || viewport.first_col > viewport.last_col) {
    return stats;
  }

  // Oversized viewport: the seed loop below visits every coordinate in the
  // rectangle, so a full-grid viewport would enumerate
  // `Sheet::kMaxRows * Sheet::kMaxCols` (~17e9) coordinates before any
  // dependency work starts. A viewport is a UI redraw region and never
  // approaches `kMaxRecalcViewportCells`; treat a larger rectangle like the
  // other invalid-viewport shapes above (no-op, dirty set preserved) so
  // callers fall back to a full `recalc()`.
  const utils::RectRange viewport_rect(viewport.first_row, viewport.first_col, viewport.last_row, viewport.last_col);
  ResourceBudget seed_budget(kMaxRecalcViewportCells);
  if (seed_budget.would_exceed(viewport_rect.size())) {
    return stats;
  }

  // As in the full pass, this label keeps the release-wave loop iterative;
  // locals in each wave's body leave scope before the backward jump.
partial_recalc_next_wave:
  // Keep the graph's spill-derived topology current before forming the
  // viewport closure. Both source additions and removals dirty their range
  // watcher; this is essential when a producer rewrite has just invalidated
  // its committed footprint.
  const DepGraph::DependencyDelta pre_dependency_delta = reconcile_spill_dependencies_locked(workbook);
  mark_edges_dirty_locked(pre_dependency_delta.added);
  mark_edges_dirty_locked(pre_dependency_delta.removed);

  // ---- Phase 1: enumerate the viewport's seed cells. ----
  // Walk the requested rectangle and pull every populated cell into the
  // seed set. Cells outside the sheet's stored extent are silently
  // dropped (the dep graph will not have entries for them anyway).
  // Phantoms of dynamic-array spills are intentionally NOT seeded
  // separately — the spill anchor is the formula cell, and resolving the
  // anchor forces the spill to refresh.
  std::vector<CellNodeId> seeds;
  seeds.reserve(static_cast<std::size_t>(viewport_rect.size()));
  for (auto [row, col] : viewport_rect) {
    // We seed every coordinate inside the viewport regardless of
    // whether it currently holds a stored cell: a viewport coordinate
    // that is presently blank may still have inbound dep-graph
    // edges (e.g. a formula on it that has been cleared but whose
    // dependents have not been re-registered yet). The closure walk
    // below tolerates absent nodes.
    seeds.push_back(CellNodeId{viewport.sheet_id, row, col});
  }

  // ---- Phase 2: compute the dependency closure backward from seeds. ----
  // The closure is "every cell whose value the viewport (transitively)
  // reads". We walk forward edges (`dependencies_of`) from each seed:
  // if A reads B, B's value contributes to A. The closure includes the
  // seed cells themselves so any dirty viewport cell still gets
  // evaluated.
  std::unordered_set<CellNodeId, CellNodeIdHash> closure;
  closure.reserve(seeds.size());
  std::vector<CellNodeId> bfs_queue = seeds;
  for (CellNodeId seed : seeds) {
    closure.insert(seed);
  }
  std::size_t bfs_head = 0;
  while (bfs_head < bfs_queue.size()) {
    const CellNodeId current = bfs_queue[bfs_head++];
    for (CellNodeId predecessor : graph_.dependencies_of_ref(current)) {
      if (closure.insert(predecessor).second) {
        bfs_queue.push_back(predecessor);
      }
    }
  }

  // Pull indexed potential spill producers into the closure: no derived edge exists before the first commit.
  bool expanded_potential_producers = true;
  std::unordered_set<std::uint32_t> processed_range_entries;
  processed_range_entries.reserve(range_dependencies_.distinct_range_count());
  Arena potential_arena;

  // The potential-producer index ignores the registry, so re-classify each candidate against it.
  const auto is_admissible_potential_producer = [&](CellNodeId producer, SpillPotential static_potential) {
    return SpillProducerIndex::admissible(workbook, registry, producer, static_potential, potential_arena);
  };

  const auto enqueue_potential_producer = [&](CellNodeId producer, SpillPotential static_potential) {
    if (!is_admissible_potential_producer(producer, static_potential)) {
      return;
    }
    if (closure.insert(producer).second) {
      expanded_potential_producers = true;
      bfs_queue.push_back(producer);
    }
  };

  // A phantom spill coordinate has no edge to its uncommitted anchor; scan for producers above-left.
  std::size_t potential_coordinate_head = 0;
  while (expanded_potential_producers) {
    expanded_potential_producers = false;
    if (dynamic.seed_pending_candidates(registry, closure, bfs_queue)) {
      expanded_potential_producers = true;
    }
    while (potential_coordinate_head < bfs_queue.size()) {
      const CellNodeId current = bfs_queue[potential_coordinate_head++];
      if (current.sheet_id >= sheet_count || !potential_spill_producers_.has_sheet(current.sheet_id)) {
        continue;
      }
      const Cell* current_cell = workbook.sheet(current.sheet_id).cell_at(current.row, current.col);
      if (blocks_spill(current_cell)) {
        continue;
      }
      const SpillTarget target(workbook.sheet(current.sheet_id), current.row, current.col);
      for (const auto& [producer, static_potential] : potential_spill_producers_.on_sheet(current.sheet_id)) {
        if (producer.sheet_id != current.sheet_id || !target.reachable_from(producer)) {
          continue;
        }
        enqueue_potential_producer(producer, static_potential);
      }
    }
    range_dependencies_.for_each_distinct_range(
        [&](std::uint32_t range_id, const CellRangeDependency& range, const std::vector<CellNodeId>& owners) {
          if (range.sheet_id >= sheet_count || !potential_spill_producers_.has_sheet(range.sheet_id)) {
            return;
          }
          if (processed_range_entries.count(range_id) != 0U) {
            return;
          }
          const bool watched_from_closure = std::any_of(
              owners.begin(), owners.end(), [&closure](CellNodeId owner) { return closure.count(owner) != 0U; });
          if (!watched_from_closure) {
            return;
          }
          processed_range_entries.insert(range_id);
          for (const auto& [producer, static_potential] : potential_spill_producers_.on_sheet(range.sheet_id)) {
            // A spill extends only down and right, so an anchor past either last bound cannot reach the range.
            if (producer.row > range.row_last || producer.col > range.col_last) {
              continue;
            }
            enqueue_potential_producer(producer, static_potential);
          }
        });
    while (bfs_head < bfs_queue.size()) {
      const CellNodeId current = bfs_queue[bfs_head++];
      for (CellNodeId predecessor : graph_.dependencies_of_ref(current)) {
        if (closure.insert(predecessor).second) {
          bfs_queue.push_back(predecessor);
          expanded_potential_producers = true;
        }
      }
    }
  }

  // ---- Phase 3: BFS-propagate dirtiness inside the closure only. ----
  // Volatile cells inside the closure are forced dirty (the caller
  // asked for fresh values within the viewport, and a volatile cell's
  // result is by definition stale). Volatile cells OUTSIDE the closure
  // are NOT touched: the user opted into a viewport-bounded recalc and
  // expects volatile cells in remote regions to wait for the next
  // full pass.
  volatiles_.for_each([&](CellNodeId cell) {
    if (closure.count(cell) == 0U) {
      return;
    }
    if (!dirty_.contains(cell)) {
      dirty_.mark(cell);
    }
    ++stats.volatile_cells;
  });

  // Now propagate dirtiness through the closure. We snapshot the dirty
  // cells that are inside the closure and BFS through reverse edges,
  // adding any newly-discovered dependents that themselves live inside
  // the closure (a dependent outside the closure is by definition
  // unreachable from the viewport, so we do not need to visit it).
  const auto for_each_dirty_in_closure = [&](auto&& visit) {
    dirty_.for_each([&](CellNodeId c) {
      if (closure.count(c) != 0U) {
        visit(c);
      }
    });
  };
  std::vector<CellNodeId> propagation_queue;
  propagation_queue.reserve(dirty_.size());
  for_each_dirty_in_closure([&](CellNodeId c) { propagation_queue.push_back(c); });
  std::size_t prop_head = 0;
  while (prop_head < propagation_queue.size()) {
    const CellNodeId current = propagation_queue[prop_head++];
    for (CellNodeId dependent : graph_.dependents_of_ref(current)) {
      if (closure.count(dependent) == 0U) {
        continue;
      }
      if (!dirty_.contains(dependent)) {
        dirty_.mark(dependent);
        propagation_queue.push_back(dependent);
      }
    }
  }

  // ---- Phase 4: Tarjan SCC + selective evaluation. ----
  // Restrict Tarjan to the dirty cells in the viewport closure. Cells
  // outside it are never visited and retain their dirty flag for a later
  // full / overlapping partial recalc.
  std::unordered_set<CellNodeId, CellNodeIdHash> dirty_closure;
  dirty_closure.reserve(propagation_queue.size());
  for_each_dirty_in_closure([&](CellNodeId c) { dirty_closure.insert(c); });
  std::vector<std::vector<CellNodeId>> sccs = condense_dirty_locked(dynamic, dirty_closure);
  std::unordered_set<CellNodeId, CellNodeIdHash> visited_in_sccs;
  RETURN_IF_ERROR(evaluate_dirty_components_locked(pass, sccs, &closure, kArenaExhaustedPartial, visited_in_sccs));

  // ---- Phase 4b: standalone dirty cells inside the closure. ----
  // Isolated formula cells with no dep-graph entries do not appear in
  // Tarjan output; sweep the closure for any dirty cell we have not
  // already touched.
  std::vector<CellNodeId> standalone_dirty;
  dirty_.for_each([&](CellNodeId c) {
    if (closure.count(c) != 0U && visited_in_sccs.count(c) == 0U) {
      standalone_dirty.push_back(c);
    }
  });
  for (CellNodeId c : standalone_dirty) {
    RETURN_IF_ERROR(evaluate_formula_cell_locked(pass, c));
  }

  // Reconcile globally after this partial wave. A newly discovered
  // watcher-to-producer edge is retried here only when that watcher belongs
  // to this viewport closure; a remote watcher remains dirty for a later
  // full or overlapping partial pass.
  const DepGraph::DependencyDelta dependency_delta = reconcile_spill_dependencies_locked(workbook);
  bool dependency_retry_in_closure = false;
  for (const DepGraph::Edge& edge : dependency_delta.added) {
    dirty_.mark(edge.first);
    if (closure.count(edge.first) != 0U) {
      dependency_retry_in_closure = true;
    }
  }
  // Judged before the closure is unmarked, so a target this viewport left
  // dirty still reads as unevaluated.
  const std::vector<CellNodeId>& stale_readers = dynamic.end_wave(&closure, &registry);
  for (const CellNodeId reader : stale_readers) {
    if (closure.count(reader) != 0U) {
      dependency_retry_in_closure = true;
    }
  }

  // ---- Phase 5: clear only the closure's dirty entries. ----
  // Cells outside the closure must remain dirty so a subsequent
  // `recalc()` (or an overlapping `partial_recalc`) revisits them.
  // Snapshot the in-closure dirty entries first to avoid mutating the
  // underlying container during iteration.
  std::vector<CellNodeId> to_unmark;
  to_unmark.reserve(closure.size());
  for_each_dirty_in_closure([&](CellNodeId c) { to_unmark.push_back(c); });
  // Hand dirtiness on to dependents outside the closure so a later pass still reaches them.
  const auto mark_outside_closure = [&](CellNodeId dependent) {
    if (closure.count(dependent) == 0U) {
      dirty_.mark(dependent);
    }
  };
  for (CellNodeId c : to_unmark) {
    for (CellNodeId dependent : graph_.dependents_of_ref(c)) {
      mark_outside_closure(dependent);
    }
    range_dependencies_.for_each_owner_covering(c, mark_outside_closure);
  }
  for (CellNodeId c : to_unmark) {
    dirty_.unmark(c);
  }
  for (const DepGraph::Edge& edge : dependency_delta.added) {
    if (closure.count(edge.first) != 0U) {
      dirty_.mark(edge.first);
    }
  }
  dynamic.mark_stale_dirty();

  // A committed spill can release a pending producer while this viewport is
  // being evaluated. Preserve releases outside the closure for a later full
  // or overlapping partial pass; releases inside the closure get a true
  // dependency-ordered next wave before this partial call returns.
  const std::vector<CellNodeId> released = release_queue.take();
  if (!released.empty() || dependency_retry_in_closure) {
    mark_spill_release_wave_locked(released);
    bool release_in_closure = false;
    for (const CellNodeId anchor : released) {
      if (closure.count(anchor) != 0U) {
        release_in_closure = true;
        break;
      }
    }
    if (release_in_closure || dependency_retry_in_closure) {
      if (!released.empty()) {
        wave_budget.count_release_wave();
        wave_budget.observe_release(workbook, released);
      }
      if (dependency_retry_in_closure) {
        wave_budget.count_dependency_wave();
      }
      if (wave_budget.exceeded(!released.empty())) {
        return make_error(FormulonErrorCode::kGraphScheduleFailed, "recalc waves made no progress",
                          "partial spill recovery or dynamic-reference retries exceeded the bounded wave budget");
      }
      goto partial_recalc_next_wave;
    }
  }
  return stats;
}

}  // namespace formulon::eval
