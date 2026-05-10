// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Returns true when the SCC `component` represents a real cycle: either
// more than one cell, or a singleton with a self-edge in the dep graph.
// Plain singletons evaluate normally; cyclic SCCs short-circuit to #REF!.
bool is_cyclic_component(const std::vector<CellNodeId>& component, const DepGraph& graph) {
  if (component.size() > 1U) {
    return true;
  }
  // Singleton: only a self-loop (cell depends on itself) is cyclic.
  const CellNodeId only = component.front();
  std::vector<CellNodeId> deps = graph.dependencies_of(only);
  return std::find(deps.begin(), deps.end(), only) != deps.end();
}

}  // namespace

// ----------------------------------------------------------------------------
// RecalcEngine
// ----------------------------------------------------------------------------

RecalcEngine::RecalcEngine() : arena_(std::make_unique<Arena>()) {}
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

void RecalcEngine::LockedMutator::unregister_formula(CellNodeId cell) const { engine_.unregister_formula_locked(cell); }

void RecalcEngine::LockedMutator::clear_cell_dependencies(CellNodeId cell) const {
  engine_.clear_cell_dependencies_locked(cell);
}

void RecalcEngine::LockedMutator::mark_dirty(CellNodeId cell) const { engine_.mark_dirty_locked(cell); }

const DepGraph& RecalcEngine::LockedMutator::dep_graph() const noexcept { return engine_.graph_; }

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
  // Same for the volatile flag — only re-register if the new AST is still
  // volatile.
  volatiles_.unregister_cell(cell);

  const ExtractedDeps deps = extract_deps(ast, cell.sheet_id, workbook);
  for (CellNodeId dep : deps.cell_deps) {
    graph_.add_dependency(cell, dep);
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
  volatiles_.unregister_cell(cell);
}

void RecalcEngine::clear_cell_dependencies(CellNodeId cell) {
  std::lock_guard<std::mutex> guard(mutex_);
  clear_cell_dependencies_locked(cell);
}

void RecalcEngine::clear_cell_dependencies_locked(CellNodeId cell) {
  graph_.clear_dependencies_of(cell);
  volatiles_.unregister_cell(cell);
}

void RecalcEngine::mark_dirty(CellNodeId cell) {
  std::lock_guard<std::mutex> guard(mutex_);
  mark_dirty_locked(cell);
}

void RecalcEngine::mark_dirty_locked(CellNodeId cell) {
  dirty_.mark(cell);
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

  // ---- Phase 3: Tarjan SCC over the entire dep graph. ----
  // Tarjan emits SCCs in reverse-topological order (leaves first), so the
  // evaluation walk below sees a cell's dependencies before the cell
  // itself.
  const std::vector<std::vector<CellNodeId>> sccs = graph_.tarjan_scc();

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
        // Either divergence (solver wrote `#NUM!` on every member),
        // iteration-limit exhaustion (solver left the last-iteration
        // values in place), or callback-driven abort (also leaves
        // last-iteration values in place). The user-visible failure
        // mode in every case is "the cycle did not resolve", so we
        // count the members in `cycle_cells` to mirror the
        // disabled-iterative-calc accounting.
        //
        // For the iteration-limit case (NOT for an aborted solve) we
        // write `#NUM!` ourselves so the cells do not retain misleading
        // partial values that the user might mistake for a converged
        // result. An aborted solve is the user's choice — leave the
        // partially-converged values intact so the UI can resume.
        if (!outcome.diverged && !outcome.aborted) {
          for (CellNodeId c : component) {
            if (c.sheet_id >= sheet_count) {
              continue;
            }
            Sheet& sheet = workbook.sheet(c.sheet_id);
            sheet.set_cell_cached_value(c.row, c.col, Value::error(ErrorCode::Num));
          }
        }
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
    sheet.set_cell_cached_value(only.row, only.col, result);
    ++stats.cells_evaluated;
  }

  // ---- Phase 4b: pick up dirty cells with no graph edges. ----
  // Isolated formula cells (e.g. `=NOW()` reading nothing) never enter the
  // dep graph, so Tarjan does not include them in any SCC. Sweep the
  // dirty set for any such cells and evaluate them as plain singletons.
  // Snapshot the set first because evaluation does not mutate `dirty_`,
  // but copying keeps the loop body trivial.
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

  // ---- Phase 1: enumerate the viewport's seed cells. ----
  // Walk the requested rectangle and pull every populated cell into the
  // seed set. Cells outside the sheet's stored extent are silently
  // dropped (the dep graph will not have entries for them anyway).
  // Phantoms of dynamic-array spills are intentionally NOT seeded
  // separately — the spill anchor is the formula cell, and resolving the
  // anchor forces the spill to refresh.
  const Sheet& view_sheet = workbook.sheet(viewport.sheet_id);
  std::vector<CellNodeId> seeds;
  seeds.reserve(static_cast<std::size_t>(viewport.last_row - viewport.first_row + 1U));
  for (std::uint32_t row = viewport.first_row; row <= viewport.last_row; ++row) {
    for (std::uint32_t col = viewport.first_col; col <= viewport.last_col; ++col) {
      // We seed every coordinate inside the viewport regardless of
      // whether it currently holds a stored cell: a viewport coordinate
      // that is presently blank may still have inbound dep-graph
      // edges (e.g. a formula on it that has been cleared but whose
      // dependents have not been re-registered yet). The closure walk
      // below tolerates absent nodes.
      (void)view_sheet;
      seeds.push_back(CellNodeId{viewport.sheet_id, row, col});
    }
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
    for (CellNodeId predecessor : graph_.dependencies_of(current)) {
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
  // We reuse the same Tarjan output as `recalc()`. Cells outside the
  // closure are skipped even if they appear in dirty SCCs, which
  // preserves their dirty flag for a future full / overlapping
  // partial recalc.
  const std::vector<std::vector<CellNodeId>> sccs = graph_.tarjan_scc();
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
      if (outcome.converged) {
        for (CellNodeId c : component) {
          (void)c;
          ++stats.cells_evaluated;
          ++stats.iterative_cells;
        }
      } else {
        if (!outcome.diverged && !outcome.aborted) {
          for (CellNodeId c : component) {
            if (c.sheet_id >= sheet_count) {
              continue;
            }
            Sheet& sheet = workbook.sheet(c.sheet_id);
            sheet.set_cell_cached_value(c.row, c.col, Value::error(ErrorCode::Num));
          }
        }
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
