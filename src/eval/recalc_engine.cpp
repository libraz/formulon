//
// Implementation of `RecalcEngine`. See `recalc_engine.h` for the public
// contract.

#include "eval/recalc_engine.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/cell_evaluator.h"
#include "eval/dep_extractor.h"
#include "eval/dep_graph.h"
#include "eval/dirty_set.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_reentry.h"
#include "eval/volatile_tracker.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/rect_iterator.h"
#include "utils/resource_budget.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
// ----------------------------------------------------------------------------
// RecalcEngine
// ----------------------------------------------------------------------------

RecalcEngine::RecalcEngine() : arena_(std::make_unique<Arena>(/*initial_chunk_bytes=*/4096, kMaxEvalArenaBytes)) {}
RecalcEngine::~RecalcEngine() = default;

// ---------------------------------------------------------------------------
// LockedMutator — passkey facade routing `Workbook`'s compound mutators
// through the engine's `*_locked` API. The facade never acquires the
// mutex; the caller (a `Workbook` member) holds `mutex_` for the entire
// scope.
// ---------------------------------------------------------------------------

void RecalcEngine::LockedMutator::register_formula(CellNodeId cell, const parser::AstNode& ast,
                                                   const Workbook& workbook) const {
  engine_.register_formula_locked(cell, ast, workbook);
}

void RecalcEngine::LockedMutator::unregister_formula(CellNodeId cell) const {
  engine_.unregister_formula_locked(cell);
}

void RecalcEngine::LockedMutator::clear_cell_dependencies(CellNodeId cell) const {
  engine_.clear_cell_dependencies_locked(cell);
}

void RecalcEngine::LockedMutator::mark_dirty(CellNodeId cell) const {
  engine_.mark_dirty_locked(cell);
}

void RecalcEngine::LockedMutator::mark_range_dependents_dirty(CellNodeId cell) const {
  engine_.mark_range_dependents_dirty_locked(cell);
}

void RecalcEngine::LockedMutator::reset_graph() const {
  engine_.reset_graph_locked();
}

const DepGraph& RecalcEngine::LockedMutator::dep_graph() const noexcept {
  return engine_.graph_;
}

// ---------------------------------------------------------------------------
// Public mutating API: each entry acquires `mutex_` and delegates to the
// `_locked` body. Internal callers (notably the parallel scheduler) take
// `mutex_` themselves and call the `_locked` helpers directly to avoid
// re-locking.
// ---------------------------------------------------------------------------

void RecalcEngine::register_formula(CellNodeId cell, const parser::AstNode& ast, const Workbook& workbook) {
  std::lock_guard<std::mutex> guard(mutex_);
  register_formula_locked(cell, ast, workbook);
}

void RecalcEngine::register_formula_locked(CellNodeId cell, const parser::AstNode& ast, const Workbook& workbook) {
  // Drop the cell's previous outgoing edges so re-registration is a clean
  // rewrite (the new dependency set may differ from the old one).
  graph_.clear_dependencies_of(cell);
  range_dependencies_.erase(
      std::remove_if(range_dependencies_.begin(), range_dependencies_.end(),
                     [cell](const RegisteredRangeDependency& entry) { return entry.dependent == cell; }),
      range_dependencies_.end());
  // Same for the volatile flag — only re-register if the new AST is still
  // volatile.
  volatiles_.unregister_cell(cell);

  const ExtractedDeps deps = extract_deps(ast, cell.sheet_id, workbook);
  for (CellNodeId dep : deps.cell_deps) {
    graph_.add_dependency(cell, dep);
  }
  for (CellRangeDependency range : deps.range_deps) {
    range_dependencies_.push_back(RegisteredRangeDependency{cell, range});

    // Preserve evaluation order for formulas already inside the range. The
    // compact range table handles literal writes (including cells created in
    // the future); explicit graph edges are needed only for formula cells so
    // Tarjan evaluates their fresh cached values before this aggregate.
    const Sheet& sheet = workbook.sheet(range.sheet_id);
    for (const auto& [row, cells] : sheet.rows()) {
      if (row < range.row_first || row > range.row_last) {
        continue;
      }
      const std::uint32_t first_col = range.col_first;
      const std::uint32_t last_col = std::min<std::uint32_t>(range.col_last, static_cast<std::uint32_t>(cells.size()));
      for (std::uint32_t col = first_col; col < last_col; ++col) {
        if (!cells[col].formula_text.empty()) {
          const CellNodeId source{range.sheet_id, row, col};
          if (source != cell) {
            graph_.add_dependency(cell, source);
          }
        }
      }
    }
  }

  // A formula added after an aggregate must also become an explicit graph
  // dependency of every existing range watcher that contains it. Otherwise
  // the watcher would be dirtied, but Tarjan would have no ordering edge to
  // ensure the new formula's cached value is refreshed first.
  for (const RegisteredRangeDependency& entry : range_dependencies_) {
    if (entry.dependent != cell && entry.range.contains(cell)) {
      graph_.add_dependency(entry.dependent, cell);
    }
  }
  if (deps.is_volatile) {
    volatiles_.register_cell(cell);
  }
}

void RecalcEngine::unregister_formula(CellNodeId cell) {
  std::lock_guard<std::mutex> guard(mutex_);
  unregister_formula_locked(cell);
}

void RecalcEngine::unregister_formula_locked(CellNodeId cell) {
  graph_.remove_node(cell);
  range_dependencies_.erase(
      std::remove_if(range_dependencies_.begin(), range_dependencies_.end(),
                     [cell](const RegisteredRangeDependency& entry) { return entry.dependent == cell; }),
      range_dependencies_.end());
  volatiles_.unregister_cell(cell);
}

void RecalcEngine::clear_cell_dependencies(CellNodeId cell) {
  std::lock_guard<std::mutex> guard(mutex_);
  clear_cell_dependencies_locked(cell);
}

void RecalcEngine::clear_cell_dependencies_locked(CellNodeId cell) {
  graph_.clear_dependencies_of(cell);
  range_dependencies_.erase(
      std::remove_if(range_dependencies_.begin(), range_dependencies_.end(),
                     [cell](const RegisteredRangeDependency& entry) { return entry.dependent == cell; }),
      range_dependencies_.end());
  volatiles_.unregister_cell(cell);
}

void RecalcEngine::mark_dirty(CellNodeId cell) {
  std::lock_guard<std::mutex> guard(mutex_);
  mark_dirty_locked(cell);
}

void RecalcEngine::mark_dirty_locked(CellNodeId cell) {
  dirty_.mark(cell);
}

void RecalcEngine::mark_range_dependents_dirty_locked(CellNodeId cell) {
  for (const RegisteredRangeDependency& entry : range_dependencies_) {
    if (entry.range.contains(cell)) {
      dirty_.mark(entry.dependent);
    }
  }
}

void RecalcEngine::reset_graph_locked() {
  graph_ = DepGraph{};
  range_dependencies_.clear();
  volatiles_.clear();
  dirty_.clear();
}

Expected<RecalcStats, Error> RecalcEngine::recalc(Workbook& workbook, const FunctionRegistry& registry) {
  // Same-thread re-entry would deadlock on the non-recursive `mutex_`.
  // The shared `g_in_recalc` flag (also tracked by `recalc_parallel_impl`)
  // converts that deadlock into a structured `kGraphRecalcReentrant`
  // error, which is the friendlier surface for a UDF or progress callback
  // that accidentally calls back into the engine.
  if (detail::g_in_recalc) {
    return make_error(FormulonErrorCode::kGraphRecalcReentrant,
                      "RecalcEngine::recalc called recursively on the same thread",
                      "the engine does not support nested recalc; the inner call is rejected");
  }
  detail::RecalcReentryGuard reentry_guard;
  std::lock_guard<std::mutex> guard(mutex_);
  return recalc_locked(workbook, registry);
}

Expected<RecalcStats, Error> RecalcEngine::recalc_locked(Workbook& workbook, const FunctionRegistry& registry) {
  RecalcStats stats;

  // ---- Phase 1: seed the dirty set with every volatile cell. ----
  // Volatile formulas re-execute every pass even if their inputs are
  // unchanged. Counting them here also reports how many of the eventually-
  // evaluated cells were forced by volatility.
  volatiles_.for_each([&](CellNodeId cell) {
    if (!dirty_.contains(cell)) {
      dirty_.mark(cell);
    }
    ++stats.volatile_cells;
  });

  // ---- Phase 2: BFS-propagate dirtiness through reverse edges. ----
  // Snapshot the seed list because `dirty_.mark` mutations during BFS
  // would otherwise shift the iteration target.
  std::vector<CellNodeId> bfs_queue;
  bfs_queue.reserve(dirty_.size());
  dirty_.for_each([&](CellNodeId c) { bfs_queue.push_back(c); });
  std::size_t bfs_head = 0;
  while (bfs_head < bfs_queue.size()) {
    const CellNodeId current = bfs_queue[bfs_head++];
    for (CellNodeId dependent : graph_.dependents_of(current)) {
      if (!dirty_.contains(dependent)) {
        dirty_.mark(dependent);
        bfs_queue.push_back(dependent);
      }
    }
  }

  // ---- Phase 3: Tarjan SCC over the dirty induced subgraph. ----
  // Tarjan emits SCCs in reverse-topological order (leaves first), so the
  // evaluation walk below sees a cell's dependencies before the cell
  // itself. Reverse-edge propagation above includes every member of a
  // reachable cycle, so excluding clean nodes preserves SCC boundaries.
  std::unordered_set<CellNodeId, CellNodeIdHash> dirty_nodes;
  dirty_nodes.reserve(dirty_.size());
  dirty_.for_each([&](CellNodeId c) { dirty_nodes.insert(c); });
  const std::vector<std::vector<CellNodeId>> sccs = graph_.tarjan_scc_subset(dirty_nodes);

  // Index sheet pointers once so we can resolve `CellNodeId::sheet_id` to
  // a `Sheet*` without a per-cell lookup. Workbook sheet count is small
  // (typically <= a few dozen).
  const std::size_t sheet_count = workbook.sheet_count();

  // ---- Phase 4: evaluate every dirty SCC. ----
  // Also track which dirty cells have been visited via the SCC walk, so a
  // final sweep can pick up dirty cells with no graph edges (e.g. a
  // standalone `=NOW()` that reads nothing). Tarjan only emits nodes that
  // appear in the forward / reverse adjacency maps; isolated formula
  // cells are absent from both.
  std::unordered_set<CellNodeId, CellNodeIdHash> visited_in_sccs;
  for (const std::vector<CellNodeId>& component : sccs) {
    // Skip components whose intersection with the dirty set is empty —
    // their cells are already up to date.
    bool any_dirty = false;
    for (CellNodeId c : component) {
      if (dirty_.contains(c)) {
        any_dirty = true;
        break;
      }
    }
    if (!any_dirty) {
      continue;
    }

    if (is_cyclic_component(component, graph_)) {
      // Track which members the dispatcher visited regardless of the
      // resolution path so the standalone-dirty sweep does not re-touch
      // them.
      for (CellNodeId c : component) {
        visited_in_sccs.insert(c);
      }

      if (!iterative_.enabled) {
        // Cycle SCC, iterative calc disabled: surface #REF! on every
        // member. Excel's analogous behaviour pops a warning dialog and
        // leaves cells at zero; Formulon collapses the diagnostic into
        // an Excel-visible error sentinel since we have no UI to host a
        // banner.
        for (CellNodeId c : component) {
          if (c.sheet_id >= sheet_count) {
            continue;  // Defensive — should not happen.
          }
          Sheet& sheet = workbook.sheet(c.sheet_id);
          sheet.set_cell_cached_value(c.row, c.col, Value::error(ErrorCode::Ref));
          ++stats.cycle_cells;
        }
        continue;
      }

      // Iterative calc enabled: hand the SCC to the solver. The
      // `evaluate_one` lambda mirrors the singleton path's evaluator
      // glue: parse on the fly, dispatch through `evaluate()`, fold
      // dynamic-array spills back into a scalar anchor. The `commit`
      // lambda writes the new value into the cell store so the next
      // iteration's `evaluate_one` reads the freshest value back.
      auto evaluate_one = [&](CellNodeId c) -> Value {
        if (c.sheet_id >= sheet_count) {
          return Value::error(ErrorCode::Ref);
        }
        Sheet& sheet = workbook.sheet(c.sheet_id);
        const Cell* cell_data = sheet.cell_at(c.row, c.col);
        if (cell_data == nullptr || cell_data->formula_text.empty()) {
          // No formula text: nothing to evaluate. Treat as Blank so the
          // solver still has a value to compare against — this happens
          // only on logic bugs but we degrade gracefully.
          return Value::blank();
        }
        // Reset the bump arena per-evaluation. The previous iteration's
        // committed Text scalars survive this reset because
        // `Sheet::set_cell_cached_value` deep-copies Text payloads into
        // each cell's own `cached_text_owned` storage; Array results are
        // similarly deep-copied into `SpillRegion::owned_strings` by
        // `Sheet::commit_spill`.
        arena_->reset();
        EvaluateCellOptions opts;
        opts.iterative_mode = true;
        return evaluate_cell_for_recalc(workbook, sheet, *cell_data, c.row, c.col, registry, *arena_, opts);
      };
      auto commit = [&](CellNodeId c, Value v) {
        if (c.sheet_id >= sheet_count) {
          return;
        }
        Sheet& sheet = workbook.sheet(c.sheet_id);
        sheet.set_cell_cached_value(c.row, c.col, v);
      };

      const IterativeOutcome outcome =
          run_iterative_solve(component, iterative_, evaluate_one, commit, progress_cb_, progress_user_data_);
      if (arena_->exhausted()) {
        return make_error(FormulonErrorCode::kOutOfMemory, "evaluation arena exhausted during iterative recalc");
      }
      if (outcome.converged) {
        // Solver wrote the converged values into the cell store; count
        // each member as evaluated (singleton-style accounting) plus
        // tagged as iterative.
        for (CellNodeId c : component) {
          (void)c;
          ++stats.cells_evaluated;
          ++stats.iterative_cells;
        }
      } else {
        // Iteration-limit exhaustion or callback-driven abort leaves the
        // last-iteration values in place. The user-visible failure mode is
        // still "the cycle did not resolve", so we
        // count the members in `cycle_cells` to mirror the
        // disabled-iterative-calc accounting.
        // A later pass can continue from the retained approximation.
        for (CellNodeId c : component) {
          (void)c;
          ++stats.cycle_cells;
        }
      }
      continue;
    }

    // Plain singleton: evaluate the cell.
    const CellNodeId only = component.front();
    visited_in_sccs.insert(only);
    if (only.sheet_id >= sheet_count) {
      continue;
    }
    Sheet& sheet = workbook.sheet(only.sheet_id);
    const Cell* cell_data = sheet.cell_at(only.row, only.col);
    if (cell_data == nullptr || cell_data->formula_text.empty()) {
      // The dep graph may carry pure-input cells (read-only literals that
      // someone reads via `add_dependency`). They have nothing to evaluate.
      continue;
    }
    // Reset the per-pass arena before each evaluate so the bump allocator
    // does not grow without bound across cells. Both result shapes
    // already survive this reset:
    //   * Array results are deep-copied into sheet-owned storage by
    //     `Sheet::commit_spill` (driven by
    //     `EvalContext::dispatch_array_result`), keeping the spill table
    //     independent of the arena.
    //   * Scalar Text results are deep-copied into the destination cell's
    //     `Cell::cached_text_owned` by `Sheet::set_cell_cached_value` on
    //     the write below, so the cached `string_view` does not dangle
    //     when the next cell's evaluation resets the arena.
    arena_->reset();
    Value result = evaluate_cell_for_recalc(workbook, sheet, *cell_data, only.row, only.col, registry, *arena_);
    if (arena_->exhausted()) {
      return make_error(FormulonErrorCode::kOutOfMemory, "evaluation arena exhausted during recalc");
    }
    sheet.set_cell_cached_value(only.row, only.col, result);
    ++stats.cells_evaluated;
  }

  // ---- Phase 4b: defensive pickup for dirty cells not visited by Tarjan. ----
  // `tarjan_scc_subset` emits isolated selected nodes, so this normally
  // remains empty. Keep the sweep as a guard for future graph mutations.
  std::vector<CellNodeId> standalone_dirty;
  dirty_.for_each([&](CellNodeId c) {
    if (visited_in_sccs.count(c) == 0U) {
      standalone_dirty.push_back(c);
    }
  });
  for (CellNodeId c : standalone_dirty) {
    if (c.sheet_id >= sheet_count) {
      continue;
    }
    Sheet& sheet = workbook.sheet(c.sheet_id);
    const Cell* cell_data = sheet.cell_at(c.row, c.col);
    if (cell_data == nullptr || cell_data->formula_text.empty()) {
      continue;
    }
    arena_->reset();
    Value result = evaluate_cell_for_recalc(workbook, sheet, *cell_data, c.row, c.col, registry, *arena_);
    if (arena_->exhausted()) {
      return make_error(FormulonErrorCode::kOutOfMemory, "evaluation arena exhausted during recalc");
    }
    sheet.set_cell_cached_value(c.row, c.col, result);
    ++stats.cells_evaluated;
  }

  // ---- Phase 5: clear the dirty set. ----
  dirty_.clear();
  return stats;
}

Expected<RecalcStats, Error> RecalcEngine::partial_recalc(Workbook& workbook, const FunctionRegistry& registry,
                                                          const SheetCellRange& viewport) {
  // Mirror `recalc()` re-entry handling: a callback that calls back into a
  // recalc API would otherwise deadlock on the engine mutex.
  if (detail::g_in_recalc) {
    return make_error(FormulonErrorCode::kGraphRecalcReentrant,
                      "RecalcEngine::partial_recalc called recursively on the same thread",
                      "the engine does not support nested recalc; the inner call is rejected");
  }
  detail::RecalcReentryGuard reentry_guard;
  std::lock_guard<std::mutex> guard(mutex_);
  return partial_recalc_locked(workbook, registry, viewport);
}

Expected<RecalcStats, Error> RecalcEngine::partial_recalc_locked(Workbook& workbook, const FunctionRegistry& registry,
                                                                 const SheetCellRange& viewport) {
  RecalcStats stats;

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

  // ---- Phase 1: enumerate the viewport's seed cells. ----
  // Walk the requested rectangle and pull every populated cell into the
  // seed set. Cells outside the sheet's stored extent are silently
  // dropped (the dep graph will not have entries for them anyway).
  // Phantoms of dynamic-array spills are intentionally NOT seeded
  // separately — the spill anchor is the formula cell, and resolving the
  // anchor forces the spill to refresh.
  const Sheet& view_sheet = workbook.sheet(viewport.sheet_id);
  std::vector<CellNodeId> seeds;
  seeds.reserve(static_cast<std::size_t>(viewport_rect.size()));
  for (auto [row, col] : viewport_rect) {
    // We seed every coordinate inside the viewport regardless of
    // whether it currently holds a stored cell: a viewport coordinate
    // that is presently blank may still have inbound dep-graph
    // edges (e.g. a formula on it that has been cleared but whose
    // dependents have not been re-registered yet). The closure walk
    // below tolerates absent nodes.
    (void)view_sheet;
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
  std::vector<CellNodeId> propagation_queue;
  propagation_queue.reserve(dirty_.size());
  dirty_.for_each([&](CellNodeId c) {
    if (closure.count(c) != 0U) {
      propagation_queue.push_back(c);
    }
  });
  std::size_t prop_head = 0;
  while (prop_head < propagation_queue.size()) {
    const CellNodeId current = propagation_queue[prop_head++];
    for (CellNodeId dependent : graph_.dependents_of(current)) {
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
  dirty_.for_each([&](CellNodeId c) {
    if (closure.count(c) != 0U) {
      dirty_closure.insert(c);
    }
  });
  const std::vector<std::vector<CellNodeId>> sccs = graph_.tarjan_scc_subset(dirty_closure);
  std::unordered_set<CellNodeId, CellNodeIdHash> visited_in_sccs;
  for (const std::vector<CellNodeId>& component : sccs) {
    // Skip components that have no overlap with the closure: their
    // cells are not transitively read by the viewport.
    bool any_in_closure = false;
    for (CellNodeId c : component) {
      if (closure.count(c) != 0U) {
        any_in_closure = true;
        break;
      }
    }
    if (!any_in_closure) {
      continue;
    }

    // Skip components that have no dirty member: nothing to do here.
    bool any_dirty = false;
    for (CellNodeId c : component) {
      if (dirty_.contains(c) && closure.count(c) != 0U) {
        any_dirty = true;
        break;
      }
    }
    if (!any_dirty) {
      continue;
    }

    if (is_cyclic_component(component, graph_)) {
      // A cycle that the viewport reaches must still be surfaced as a
      // cycle: the closure restriction never hides a circular reference.
      // Mirrors `recalc()`'s cycle handling exactly, with the iterative
      // solver wired to the same progress callback.
      for (CellNodeId c : component) {
        visited_in_sccs.insert(c);
      }

      if (!iterative_.enabled) {
        for (CellNodeId c : component) {
          if (c.sheet_id >= sheet_count) {
            continue;
          }
          Sheet& sheet = workbook.sheet(c.sheet_id);
          sheet.set_cell_cached_value(c.row, c.col, Value::error(ErrorCode::Ref));
          ++stats.cycle_cells;
        }
        continue;
      }

      auto evaluate_one = [&](CellNodeId c) -> Value {
        if (c.sheet_id >= sheet_count) {
          return Value::error(ErrorCode::Ref);
        }
        Sheet& sheet = workbook.sheet(c.sheet_id);
        const Cell* cell_data = sheet.cell_at(c.row, c.col);
        if (cell_data == nullptr || cell_data->formula_text.empty()) {
          return Value::blank();
        }
        arena_->reset();
        EvaluateCellOptions opts;
        opts.iterative_mode = true;
        return evaluate_cell_for_recalc(workbook, sheet, *cell_data, c.row, c.col, registry, *arena_, opts);
      };
      auto commit = [&](CellNodeId c, Value v) {
        if (c.sheet_id >= sheet_count) {
          return;
        }
        Sheet& sheet = workbook.sheet(c.sheet_id);
        sheet.set_cell_cached_value(c.row, c.col, v);
      };

      const IterativeOutcome outcome =
          run_iterative_solve(component, iterative_, evaluate_one, commit, progress_cb_, progress_user_data_);
      if (arena_->exhausted()) {
        return make_error(FormulonErrorCode::kOutOfMemory, "evaluation arena exhausted during partial recalc");
      }
      if (outcome.converged) {
        for (CellNodeId c : component) {
          (void)c;
          ++stats.cells_evaluated;
          ++stats.iterative_cells;
        }
      } else {
        // As in full recalc, only an actual divergence writes `#NUM!`.
        // A finite iteration budget leaves the solver's final approximation
        // in place so a later viewport/full pass can resume from it.
        for (CellNodeId c : component) {
          (void)c;
          ++stats.cycle_cells;
        }
      }
      continue;
    }

    // Plain singleton: only evaluate if the cell is in the closure AND
    // dirty. Cells outside the closure stay dirty for a later pass.
    const CellNodeId only = component.front();
    if (closure.count(only) == 0U || !dirty_.contains(only)) {
      continue;
    }
    visited_in_sccs.insert(only);
    if (only.sheet_id >= sheet_count) {
      continue;
    }
    Sheet& sheet = workbook.sheet(only.sheet_id);
    const Cell* cell_data = sheet.cell_at(only.row, only.col);
    if (cell_data == nullptr || cell_data->formula_text.empty()) {
      continue;
    }
    arena_->reset();
    Value result = evaluate_cell_for_recalc(workbook, sheet, *cell_data, only.row, only.col, registry, *arena_);
    if (arena_->exhausted()) {
      return make_error(FormulonErrorCode::kOutOfMemory, "evaluation arena exhausted during partial recalc");
    }
    sheet.set_cell_cached_value(only.row, only.col, result);
    ++stats.cells_evaluated;
  }

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
    if (c.sheet_id >= sheet_count) {
      continue;
    }
    Sheet& sheet = workbook.sheet(c.sheet_id);
    const Cell* cell_data = sheet.cell_at(c.row, c.col);
    if (cell_data == nullptr || cell_data->formula_text.empty()) {
      continue;
    }
    arena_->reset();
    Value result = evaluate_cell_for_recalc(workbook, sheet, *cell_data, c.row, c.col, registry, *arena_);
    if (arena_->exhausted()) {
      return make_error(FormulonErrorCode::kOutOfMemory, "evaluation arena exhausted during partial recalc");
    }
    sheet.set_cell_cached_value(c.row, c.col, result);
    ++stats.cells_evaluated;
  }

  // ---- Phase 5: clear only the closure's dirty entries. ----
  // Cells outside the closure must remain dirty so a subsequent
  // `recalc()` (or an overlapping `partial_recalc`) revisits them.
  // Snapshot the in-closure dirty entries first to avoid mutating the
  // underlying container during iteration.
  std::vector<CellNodeId> to_unmark;
  to_unmark.reserve(closure.size());
  dirty_.for_each([&](CellNodeId c) {
    if (closure.count(c) != 0U) {
      to_unmark.push_back(c);
    }
  });
  for (CellNodeId c : to_unmark) {
    dirty_.unmark(c);
  }
  return stats;
}

}  // namespace formulon::eval
