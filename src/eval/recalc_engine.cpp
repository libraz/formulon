//
// Implementation of `RecalcEngine`. See `recalc_engine.h` for the public
// contract.

#include "eval/recalc_engine.h"

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

// Same-thread re-entry would deadlock on the non-recursive engine mutex. The
// shared `g_in_recalc` flag (also tracked by `recalc_parallel_impl`) turns
// that into a structured `kGraphRecalcReentrant` error, the friendlier
// surface for a UDF or progress callback that calls back into the engine.
std::optional<Error> reentry_error(const char* what) {
  if (!detail::g_in_recalc) {
    return std::nullopt;
  }
  return make_error(FormulonErrorCode::kGraphRecalcReentrant,
                    std::string("RecalcEngine::") + what + " called recursively on the same thread",
                    "the engine does not support nested recalc; the inner call is rejected");
}

const char* const kArenaExhaustedRecalc = "evaluation arena exhausted during recalc";
const char* const kArenaExhaustedIterative = "evaluation arena exhausted during iterative recalc";

// The spill-release queue, the progress snapshot and the wave ceilings are
// shared with the parallel scheduler; see `spill_release.h`.
using detail::queue_spill_release;
using detail::SpillReleaseQueue;
using detail::SpillWaveBudget;

// Drops the virtual node of every compact rectangle that lost its last watcher.
void remove_range_nodes(DepGraph& graph, const std::vector<std::uint32_t>& released_range_ids) {
  for (const std::uint32_t range_id : released_range_ids) {
    graph.remove_node(range_node(range_id));
  }
}

}  // namespace

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

void RecalcEngine::LockedMutator::mark_range_dependents_dirty(const CellRangeDependency& range) const {
  engine_.mark_range_dependents_dirty_locked(range);
}

void RecalcEngine::LockedMutator::reset_graph() const {
  engine_.reset_graph_locked();
}

const DepGraph& RecalcEngine::LockedMutator::dep_graph() const noexcept {
  return engine_.graph_;
}

bool RecalcEngine::LockedMutator::has_ever_registered_formula() const noexcept {
  return engine_.has_ever_registered_formula_;
}

std::vector<CellNodeId> RecalcEngine::LockedMutator::three_d_span_owners_covering_sheet(
    std::uint16_t edited_sheet) const {
  std::vector<CellNodeId> owners;
  owners.reserve(engine_.three_d_span_dependencies_.size());
  std::unordered_set<CellNodeId, CellNodeIdHash> seen;
  seen.reserve(engine_.three_d_span_dependencies_.size());
  for (const RegisteredThreeDSpan& entry : engine_.three_d_span_dependencies_) {
    if (edited_sheet < entry.span.sheet_first || edited_sheet > entry.span.sheet_last) {
      continue;
    }
    if (seen.insert(entry.owner).second) {
      owners.push_back(entry.owner);
    }
  }
  std::sort(owners.begin(), owners.end(), CellNodeIdOrder{});
  return owners;
}

void RecalcEngine::ReferencedCellIndex::add(CellNodeId cell) {
  if (cell.sheet_id >= by_sheet_.size()) {
    by_sheet_.resize(static_cast<std::size_t>(cell.sheet_id) + 1U);
  }
  ++by_sheet_[cell.sheet_id][key(cell.row, cell.col)];
}

void RecalcEngine::ReferencedCellIndex::remove(CellNodeId cell) {
  if (cell.sheet_id >= by_sheet_.size()) {
    return;
  }
  std::map<std::uint64_t, std::uint32_t>& cells = by_sheet_[cell.sheet_id];
  const auto found = cells.find(key(cell.row, cell.col));
  if (found != cells.end() && --found->second == 0U) {
    cells.erase(found);
  }
}

bool RecalcEngine::ReferencedCellIndex::empty() const noexcept {
  return std::all_of(by_sheet_.begin(), by_sheet_.end(), [](const auto& cells) { return cells.empty(); });
}

void RecalcEngine::forget_referenced_cells_locked(CellNodeId cell) {
  for (const CellNodeId dep : graph_.dependencies_of_ref(cell)) {
    if (!is_range_node(dep) && graph_.has_dependency_source(cell, dep, DepGraph::DependencySource::kAuthored)) {
      referenced_cells_.remove(dep);
    }
  }
}

void RecalcEngine::mark_edges_dirty_locked(const std::vector<DepGraph::Edge>& edges) {
  for (const DepGraph::Edge& edge : edges) {
    dirty_.mark(edge.first);
  }
}

void RecalcEngine::seed_dirty_locked(std::uint32_t* volatile_cells) {
  volatiles_.for_each([&](CellNodeId cell) {
    if (!dirty_.contains(cell)) {
      dirty_.mark(cell);
    }
    if (volatile_cells != nullptr) {
      ++*volatile_cells;
    }
  });

  // Snapshot the seed list because `dirty_.mark` mutations during BFS
  // would otherwise shift the iteration target.
  std::vector<CellNodeId> bfs_queue;
  bfs_queue.reserve(dirty_.size());
  dirty_.for_each([&](CellNodeId c) { bfs_queue.push_back(c); });
  std::size_t bfs_head = 0;
  while (bfs_head < bfs_queue.size()) {
    const CellNodeId current = bfs_queue[bfs_head++];
    for (CellNodeId dependent : graph_.dependents_of_ref(current)) {
      if (!dirty_.contains(dependent)) {
        dirty_.mark(dependent);
        bfs_queue.push_back(dependent);
      }
    }
  }
}

Expected<void, Error> RecalcEngine::evaluate_dirty_components_locked(
    const SerialEvalPass& pass, const std::vector<std::vector<CellNodeId>>& sccs,
    const std::unordered_set<CellNodeId, CellNodeIdHash>* closure, const char* iterative_oom_message,
    std::unordered_set<CellNodeId, CellNodeIdHash>& visited) {
  for (const std::vector<CellNodeId>& component : sccs) {
    // Skip components with no dirty member (inside the closure, when restricted): their cells are up to date or are
    // left dirty for a later pass.
    const bool any_dirty = std::any_of(component.begin(), component.end(), [&](CellNodeId c) {
      return dirty_.contains(c) && (closure == nullptr || closure->count(c) != 0U);
    });
    if (!any_dirty) {
      continue;
    }

    if (is_cyclic_component(component, graph_)) {
      // A cycle the walk reaches is always surfaced as a cycle. Members are recorded as visited on every resolution
      // path so the standalone-dirty sweep does not re-touch them.
      for (CellNodeId c : component) {
        visited.insert(c);
      }
      RETURN_IF_ERROR(evaluate_cyclic_component_locked(pass, component, iterative_oom_message));
      continue;
    }

    const CellNodeId only = component.front();
    visited.insert(only);
    RETURN_IF_ERROR(evaluate_formula_cell_locked(pass, only));
  }
  return Expected<void, Error>::Ok();
}

void RecalcEngine::commit_unresolved_cycle_locked(const SerialEvalPass& pass,
                                                  const std::vector<CellNodeId>& component) {
  // Excel shows a warning and leaves the cells as they were; with no UI to
  // host that banner Formulon surfaces #REF! (0 for a cell that is its own
  // cycle), keeping Excel's last value only for a cycle through an OFFSET /
  // INDIRECT read.
  const bool dynamic_cycle = closes_through_dynamic_edge(component, graph_);
  const Value unresolved = unresolved_cycle_value(component);
  const std::uint64_t ordinal = pass.dynamic.next_ordinal();
  if (dynamic_cycle) {
    pass.dynamic.log().restore_prior_values(pass.workbook, component);
  }
  for (const CellNodeId c : component) {
    if (c.sheet_id >= pass.workbook.sheet_count()) {
      continue;  // A virtual range node.
    }
    if (!dynamic_cycle) {
      Sheet& sheet = pass.workbook.sheet(c.sheet_id);
      SpillCommitter committer(&sheet, c.row, c.col, pass.release_callback, pass.release_user_data);
      sheet.set_cell_cached_value(c.row, c.col, committer.commit(unresolved));
    }
    pass.dynamic.log().note_commit(c, ordinal);
    ++pass.stats.cycle_cells;
  }
}

Expected<void, Error> RecalcEngine::evaluate_formula_cell_locked(const SerialEvalPass& pass, CellNodeId cell) {
  if (cell.sheet_id >= pass.workbook.sheet_count()) {
    return Expected<void, Error>::Ok();
  }
  Sheet& sheet = pass.workbook.sheet(cell.sheet_id);
  Cell staged;
  if (!stage_formula_cell(sheet, cell.row, cell.col, staged)) {
    // The dep graph may carry pure-input cells (read-only literals that
    // someone reads via `add_dependency`). They have nothing to evaluate.
    return Expected<void, Error>::Ok();
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
  EvaluateCellOptions opts;
  opts.spill_release_callback = pass.release_callback;
  opts.spill_release_user_data = pass.release_user_data;
  const std::uint64_t ordinal = pass.dynamic.next_ordinal();
  pass.dynamic.log().observe(cell, ordinal, opts, &sheet);
  Value result =
      evaluate_cell_for_recalc(pass.workbook, sheet, staged, cell.row, cell.col, pass.registry, *arena_, opts);
  pass.dynamic.log().end();
  if (arena_->exhausted()) {
    return make_error(FormulonErrorCode::kOutOfMemory, pass.oom_message);
  }
  sheet.set_cell_cached_value(cell.row, cell.col, result);
  pass.dynamic.log().note_commit(cell, ordinal);
  ++pass.stats.cells_evaluated;
  return Expected<void, Error>::Ok();
}

Expected<void, Error> RecalcEngine::evaluate_cyclic_component_locked(const SerialEvalPass& pass,
                                                                     const std::vector<CellNodeId>& component,
                                                                     const char* iterative_oom_message) {
  CyclicComponentEvaluator evaluator(pass.workbook, pass.registry, *arena_, pass.release_callback,
                                     pass.release_user_data, &pass.dynamic.log(), pass.dynamic.next_ordinal(),
                                     /*commit_mutex=*/nullptr);
  auto evaluate_one = [&](CellNodeId c) { return evaluator.evaluate(c, nullptr); };
  auto commit = [&](CellNodeId c, Value v) { evaluator.commit(c, v); };

  const std::vector<CellNodeId> cells = cells_of_component(component);
  if (!iterative_.enabled) {
    // Excel judges circularity by what evaluation reads, so only a cycle
    // that reads back into itself is one. A cycle through an OFFSET /
    // INDIRECT read keeps its own treatment.
    const bool settled = !closes_through_dynamic_edge(component, graph_) && settle_component_by_reads(cells, evaluator);
    if (arena_->exhausted()) {
      return make_error(FormulonErrorCode::kOutOfMemory, pass.oom_message);
    }
    if (settled) {
      pass.stats.cells_evaluated += static_cast<std::uint32_t>(cells.size());
      return Expected<void, Error>::Ok();
    }
    commit_unresolved_cycle_locked(pass, component);
    return Expected<void, Error>::Ok();
  }
  // A reader first evaluated as a singleton in this recalc has already
  // stepped once; the solver starts from the value it showed before.
  pass.dynamic.log().restore_prior_values(pass.workbook, cells);
  prepare_iterative_component_seeds(pass.workbook, cells);
  const IterativeOutcome outcome =
      run_iterative_solve(cells, iterative_, evaluate_one, commit, progress_cb_, progress_user_data_);
  if (arena_->exhausted()) {
    return make_error(FormulonErrorCode::kOutOfMemory, iterative_oom_message);
  }
  if (outcome.converged) {
    // Solver wrote the converged values into the cell store; count
    // each member as evaluated (singleton-style accounting) plus
    // tagged as iterative.
    pass.stats.cells_evaluated += static_cast<std::uint32_t>(cells.size());
    pass.stats.iterative_cells += static_cast<std::uint32_t>(cells.size());
  } else {
    // Iteration-limit exhaustion or callback-driven abort leaves the
    // last-iteration values in place. The user-visible failure mode is
    // still "the cycle did not resolve", so we
    // count the members in `cycle_cells` to mirror the
    // disabled-iterative-calc accounting.
    // A later pass can continue from the retained approximation.
    pass.stats.cycle_cells += static_cast<std::uint32_t>(cells.size());
  }
  return Expected<void, Error>::Ok();
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
  // Keep this separate from graph node presence: a syntactically valid
  // formula that currently references a missing sheet can have no static
  // dependency edges at all, yet must be reconsidered when that sheet is
  // appended or renamed into existence.
  has_ever_registered_formula_ = true;

  // Drop the cell's previous outgoing edges and registrations (including its
  // spill-candidate membership, since a rewrite can change the result shape)
  // so re-registration is a clean rewrite.
  forget_referenced_cells_locked(cell);
  graph_.clear_dependencies_of(cell);
  drop_cell_registrations_locked(cell);

  const ExtractedDeps deps = extract_deps(ast, cell, workbook);
  for (CellNodeId dep : deps.cell_deps) {
    graph_.add_dependency(cell, dep);
    referenced_cells_.add(dep);
  }
  for (const ThreeDSheetSpanDependency span : deps.three_d_spans) {
    three_d_span_dependencies_.push_back(RegisteredThreeDSpan{cell, span});
    three_d_span_owners_.insert(cell);
  }
  for (const CellRangeDependency& range : deps.range_deps) {
    // The watcher reads the rectangle's virtual node, which every watcher of
    // the same rectangle shares.
    const RangeDepIndex::AddResult added = range_dependencies_.add(cell, range);
    const CellNodeId node = range_node(added.range_id);
    graph_.add_dependency(cell, node);
    if (!added.is_new) {
      continue;
    }

    // Preserve evaluation order for formulas already inside the range. The
    // compact range table handles literal writes (including cells created in
    // the future); explicit graph edges are needed only for formula cells so
    // Tarjan evaluates their fresh cached values before this aggregate.
    //
    // The sheet's formula-cell index answers that directly, so the cost is
    // the formula cells inside the rectangle — never its area, and never the
    // literal content a whole-column reference spans. A watcher inside its
    // own rectangle reads itself, which is a cycle as in the flattened path.
    const Sheet& sheet = workbook.sheet(range.sheet_id);
    for (const CellAddress source :
         sheet.formula_cells_in(range.row_first, range.col_first, range.row_last, range.col_last)) {
      graph_.add_dependency(node, CellNodeId{range.sheet_id, source.row, source.col});
    }
  }

  // A formula added after an aggregate must also become an explicit graph
  // dependency of every existing rectangle that contains it. Otherwise the
  // watcher would be dirtied, but Tarjan would have no ordering edge to
  // ensure the new formula's cached value is refreshed first.
  range_dependencies_.for_each_range_covering(
      cell, [this, cell](std::uint32_t range_id) { graph_.add_dependency(range_node(range_id), cell); });
  if (deps.is_volatile) {
    volatiles_.register_cell(cell, deps.has_dynamic_reference ? VolatileKind::kDynamicReference : VolatileKind::kValue);
  }
  potential_spill_producers_.update(cell, spill_potential(ast));
}

void RecalcEngine::unregister_formula(CellNodeId cell) {
  std::lock_guard<std::mutex> guard(mutex_);
  unregister_formula_locked(cell);
}

void RecalcEngine::unregister_formula_locked(CellNodeId cell) {
  // Preserve the pre-unregistration reverse snapshot. Removing the node
  // would otherwise erase the only path that can wake compact range
  // watchers (including a watcher that was linked through a previous spill
  // footprint), leaving them with a stale cached aggregate.
  for (CellNodeId dependent : graph_.dependents_of_ref(cell)) {
    dirty_.mark(dependent);
  }
  mark_range_dependents_dirty_locked(cell);
  forget_referenced_cells_locked(cell);
  graph_.remove_node(cell);
  drop_cell_registrations_locked(cell);
}

void RecalcEngine::clear_cell_dependencies(CellNodeId cell) {
  std::lock_guard<std::mutex> guard(mutex_);
  clear_cell_dependencies_locked(cell);
}

void RecalcEngine::clear_cell_dependencies_locked(CellNodeId cell) {
  forget_referenced_cells_locked(cell);
  graph_.clear_dependencies_of(cell);
  drop_cell_registrations_locked(cell);
}

void RecalcEngine::drop_cell_registrations_locked(CellNodeId cell) {
  remove_range_nodes(graph_, range_dependencies_.erase_owner(cell));
  if (three_d_span_owners_.erase(cell) != 0U) {
    three_d_span_dependencies_.erase(
        std::remove_if(three_d_span_dependencies_.begin(), three_d_span_dependencies_.end(),
                       [cell](const RegisteredThreeDSpan& entry) { return entry.owner == cell; }),
        three_d_span_dependencies_.end());
  }
  volatiles_.unregister_cell(cell);
  potential_spill_producers_.update(cell, SpillPotential::kNever);
}

void RecalcEngine::mark_dirty(CellNodeId cell) {
  std::lock_guard<std::mutex> guard(mutex_);
  mark_dirty_locked(cell);
}

void RecalcEngine::mark_dirty_locked(CellNodeId cell) {
  dirty_.mark(cell);
}

void RecalcEngine::mark_range_dependents_dirty_locked(CellNodeId cell) {
  range_dependencies_.for_each_new_owner_covering(cell, dirty_.generation(),
                                                  [this](CellNodeId owner) { dirty_.mark(owner); });
}

void RecalcEngine::mark_range_dependents_dirty_locked(const CellRangeDependency& range) {
  range_dependencies_.for_each_owner_intersecting(range, [this](CellNodeId owner) { dirty_.mark(owner); });
}

void RecalcEngine::reset_graph_locked() {
  graph_ = DepGraph{};
  range_dependencies_.clear();
  referenced_cells_.clear();
  three_d_span_dependencies_.clear();
  three_d_span_owners_.clear();
  potential_spill_producers_.clear();
  volatiles_.clear();
  dirty_.clear();
}

Expected<RecalcStats, Error> RecalcEngine::recalc(Workbook& workbook, const FunctionRegistry& registry) {
  if (std::optional<Error> reentry = reentry_error("recalc")) {
    return *reentry;
  }
  detail::RecalcReentryGuard reentry_guard;
  std::lock_guard<std::mutex> guard(mutex_);
  return recalc_locked(workbook, registry);
}

Expected<RecalcStats, Error> RecalcEngine::recalc_locked(Workbook& workbook, const FunctionRegistry& registry) {
  RecalcStats stats;
  SpillReleaseQueue release_queue{&workbook};
  const SpillReleaseCallback release_callback = &queue_spill_release;
  SpillWaveBudget wave_budget;
  DynamicReadPass dynamic(*this, workbook);
  const SerialEvalPass pass{
      workbook, registry, dynamic, release_callback, &release_queue, stats, kArenaExhaustedRecalc};

  // Per-wave locals begin after this label and are destroyed on the backward
  // jump, while the counters, queue, and accumulated stats above persist.
recalc_next_wave:
  // Reconcile before dirtiness propagation and SCC construction. A formula
  // rewrite can clear a committed spill before this wave; removing the old
  // derived edge here prevents a stale phantom-only cycle, while additions
  // wake the watcher before topology is built.
  const DepGraph::DependencyDelta pre_dependency_delta = reconcile_spill_dependencies_locked(workbook);
  mark_edges_dirty_locked(pre_dependency_delta.added);
  mark_edges_dirty_locked(pre_dependency_delta.removed);

  // ---- Phases 1-2: seed volatile cells, then BFS through reverse edges. ----
  // Volatile formulas re-execute every pass even if their inputs are
  // unchanged. Counting them here also reports how many of the eventually-
  // evaluated cells were forced by volatility.
  seed_dirty_locked(&stats.volatile_cells);

  // ---- Phase 3: Tarjan SCC over the dirty induced subgraph. ----
  // Tarjan emits SCCs in reverse-topological order (leaves first), so the
  // evaluation walk below sees a cell's dependencies before the cell
  // itself. Reverse-edge propagation above includes every member of a
  // reachable cycle, so excluding clean nodes preserves SCC boundaries.
  std::unordered_set<CellNodeId, CellNodeIdHash> dirty_nodes;
  dirty_nodes.reserve(dirty_.size());
  dirty_.for_each([&](CellNodeId c) { dirty_nodes.insert(c); });
  std::vector<std::vector<CellNodeId>> sccs = graph_.tarjan_scc_subset(dirty_nodes);
  dynamic.settle_sccs(sccs, dirty_nodes);
  if (iterative_.enabled) {
    fuse_cyclic_components(&sccs, graph_);
  }

  // ---- Phase 4: evaluate every dirty SCC. ----
  // Also track which dirty cells have been visited via the SCC walk, so a
  // final sweep can pick up dirty cells with no graph edges (e.g. a
  // standalone `=NOW()` that reads nothing). Tarjan only emits nodes that
  // appear in the forward / reverse adjacency maps; isolated formula
  // cells are absent from both.
  std::unordered_set<CellNodeId, CellNodeIdHash> visited_in_sccs;
  RETURN_IF_ERROR(evaluate_dirty_components_locked(pass, sccs, nullptr, kArenaExhaustedIterative, visited_in_sccs));

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
    RETURN_IF_ERROR(evaluate_formula_cell_locked(pass, c));
  }

  // Reconcile compact-range -> spill-producer edges only after every cell in
  // this wave has committed its result. Newly acquired source edges schedule
  // their watcher for one targeted follow-up wave; old edges remain active
  // throughout the current wave, so a producer shrink still dirties and
  // orders its watcher correctly.
  const DepGraph::DependencyDelta dependency_delta = reconcile_spill_dependencies_locked(workbook);
  // A reader that saw a target before its final value (or itself) runs
  // again behind the edge it just taught the graph.
  const bool dynamic_retry = !dynamic.end_wave(nullptr).empty();
  const bool dependency_retry = !dependency_delta.added.empty() || dynamic_retry;
  if (dependency_retry) {
    wave_budget.count_dependency_wave();
    mark_edges_dirty_locked(dependency_delta.added);
  }

  // A spill shape change can release cells that blocked a different pending
  // producer. Keep those anchors for a dependency-ordered next wave instead
  // of losing them when this pass clears its dirty set.
  const std::vector<CellNodeId> released = release_queue.take();
  if (!released.empty() || dependency_retry) {
    wave_budget.count_release_wave();
    wave_budget.observe_release(workbook, released);
    if (wave_budget.exceeded(!released.empty())) {
      // Do not recurse through an unbounded chain of release waves. Preserve
      // the existing dirty set (including unrelated work) and keep the
      // release targets dirty for a caller retry after an external mutation.
      mark_spill_release_wave_locked(released);
      return make_error(FormulonErrorCode::kGraphScheduleFailed, "recalc waves made no progress",
                        "spill recovery or dynamic-reference retries exceeded the bounded wave budget");
    }
    dirty_.clear();
    mark_spill_release_wave_locked(released);
    mark_edges_dirty_locked(dependency_delta.added);
    dynamic.mark_stale_dirty();
    goto recalc_next_wave;
  }

  // ---- Phase 5: clear the dirty set. ----
  dirty_.clear();
  disabled_cycle_refs_pending_ = false;
  return stats;
}

Expected<RecalcStats, Error> RecalcEngine::partial_recalc(Workbook& workbook, const FunctionRegistry& registry,
                                                          const SheetCellRange& viewport) {
  if (std::optional<Error> reentry = reentry_error("partial_recalc")) {
    return *reentry;
  }
  detail::RecalcReentryGuard reentry_guard;
  std::lock_guard<std::mutex> guard(mutex_);
  return partial_recalc_locked(workbook, registry, viewport);
}

}  // namespace formulon::eval
