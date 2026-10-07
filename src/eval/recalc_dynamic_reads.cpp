//
// `RecalcEngine::DynamicReadPass` and the dynamic-only cycle drop: the runtime-read
// reconciliation of a recalculation pass. See `recalc_engine.h`.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <unordered_set>
#include <vector>

#include "eval/dep_graph.h"
#include "eval/dynamic_read_log.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "sheet.h"
#include "utils/arena.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Largest runtime read rectangle whose open cells are checked one by one before admitting a producer.
constexpr std::uint64_t kMaxCheckedSpillReadCells = 256U;

}  // namespace

RecalcEngine::DynamicReadPass::DynamicReadPass(RecalcEngine& engine, const Workbook& workbook)
    : engine_(engine), workbook_(workbook), log_(workbook, engine.volatiles_, engine.graph_) {}

void RecalcEngine::DynamicReadPass::settle_sccs(std::vector<std::vector<CellNodeId>>& sccs,
                                                const std::unordered_set<CellNodeId, CellNodeIdHash>& nodes) {
  if (!first_wave_) {
    return;
  }
  // Before anything is evaluated: a cycle through a dynamic read is left at
  // these values whichever path, serial or pooled, evaluates its members.
  log_.keep_prior_values_fed_by_readers(nodes);
  if (engine_.drop_dynamic_only_cycles_locked(sccs)) {
    sccs = engine_.graph_.tarjan_scc_subset(nodes);
  }
}

const std::vector<CellNodeId>& RecalcEngine::DynamicReadPass::end_wave(
    const std::unordered_set<CellNodeId, CellNodeIdHash>* closure, const FunctionRegistry* registry) {
  stale_ = engine_.reconcile_dynamic_reads_locked(workbook_, log_, closure, refreshed_);
  if (closure != nullptr && registry != nullptr) {
    discover_pending_candidates(*registry, *closure);
  }
  log_.clear_wave();
  first_wave_ = false;
  std::sort(stale_.begin(), stale_.end(), CellNodeIdOrder{});
  stale_.erase(std::unique(stale_.begin(), stale_.end()), stale_.end());
  return stale_;
}

void RecalcEngine::DynamicReadPass::discover_pending_candidates(
    const FunctionRegistry& registry, const std::unordered_set<CellNodeId, CellNodeIdHash>& closure) {
  Arena potential_arena;
  for (const DynamicRead& read : log_.reads()) {
    if (read.sheet_id >= workbook_.sheet_count() || !engine_.potential_spill_producers_.has_sheet(read.sheet_id)) {
      continue;
    }

    // Discovery is only for a blank coordinate whose producer has not committed. A small rectangle is checked cell
    // by cell; a larger one keeps the geometric test alone.
    const Sheet& read_sheet = workbook_.sheet(static_cast<std::uint16_t>(read.sheet_id));
    const bool checks_targets =
        static_cast<std::uint64_t>(read.rect.rows()) * read.rect.cols() <= kMaxCheckedSpillReadCells;
    std::vector<SpillTarget> open_targets;
    if (checks_targets) {
      for (std::uint32_t row = read.rect.row_first; row <= read.rect.row_last; ++row) {
        for (std::uint32_t col = read.rect.col_first; col <= read.rect.col_last; ++col) {
          Sheet::CellRead target;
          read_sheet.read_formula_cell(row, col, target);
          if (!target.is_formula() && target.value().is_blank()) {
            open_targets.emplace_back(read_sheet, row, col);
          }
        }
      }
      if (open_targets.empty()) {
        continue;
      }
    }

    CandidateSet* pending = nullptr;
    CandidateSet* attempted = nullptr;
    const auto attempted_it = attempted_candidates_.find(read.reader);
    if (attempted_it != attempted_candidates_.end()) {
      attempted = &attempted_it->second;
    }
    for (const auto& [producer, static_potential] : engine_.potential_spill_producers_.on_sheet(read.sheet_id)) {
      // A spill extends only down and right, so an anchor past either last coordinate cannot intersect the read.
      if (producer.row > read.rect.row_last || producer.col > read.rect.col_last) {
        continue;
      }
      const CellNodeId candidate = producer;
      if (checks_targets && std::none_of(open_targets.begin(), open_targets.end(),
                                         [candidate](const SpillTarget& t) { return t.reachable_from(candidate); })) {
        continue;
      }
      if (log_.commit_ordinal(producer).has_value()) {
        // A producer committed in this wave is covered by spill-footprint reconciliation.
        continue;
      }
      if (closure.count(producer) != 0U) {
        // A producer already in the closure was expanded by the ordinary BFS.
        continue;
      }
      if (attempted != nullptr && attempted->count(producer) != 0U) {
        continue;
      }
      if (!SpillProducerIndex::admissible(workbook_, registry, producer, static_potential, potential_arena)) {
        continue;
      }
      if (pending == nullptr) {
        pending = &pending_candidates_[read.reader];
      }
      if (pending->insert(producer).second) {
        stale_.push_back(read.reader);
      }
    }
    // Drop an empty entry so the next wave does not visit it.
    if (pending != nullptr && pending->empty()) {
      pending_candidates_.erase(read.reader);
    }
  }
}

bool RecalcEngine::DynamicReadPass::seed_pending_candidates(const FunctionRegistry& registry,
                                                            std::unordered_set<CellNodeId, CellNodeIdHash>& closure,
                                                            std::vector<CellNodeId>& bfs_queue) {
  bool expanded = false;
  Arena potential_arena;
  for (auto pending_it = pending_candidates_.begin(); pending_it != pending_candidates_.end();) {
    const CellNodeId reader = pending_it->first;
    if (closure.count(reader) == 0U) {
      ++pending_it;
      continue;
    }
    CandidateSet& attempted = attempted_candidates_[reader];
    for (const CellNodeId producer : pending_it->second) {
      attempted.insert(producer);
      const SpillPotential* candidate_potential = engine_.potential_spill_producers_.find(producer);
      if (candidate_potential == nullptr) {
        continue;
      }
      if (!SpillProducerIndex::admissible(workbook_, registry, producer, *candidate_potential, potential_arena)) {
        continue;
      }
      if (closure.insert(producer).second) {
        bfs_queue.push_back(producer);
        expanded = true;
      }
    }
    pending_it = pending_candidates_.erase(pending_it);
  }

  // An attempted pair remains a temporary seed for every later wave of this
  // partial call. A producer can be clean when first admitted, then become
  // dirty after a compact-range watcher learns its spill footprint. Keeping
  // the seed lets the next closure expand through that producer's authored
  // dependencies without rediscovering the pair or adding a speculative
  // graph edge. Revalidation removes candidates that were rewritten away.
  for (auto attempted_it = attempted_candidates_.begin(); attempted_it != attempted_candidates_.end();) {
    if (closure.count(attempted_it->first) == 0U) {
      ++attempted_it;
      continue;
    }
    CandidateSet& attempted = attempted_it->second;
    for (auto producer_it = attempted.begin(); producer_it != attempted.end();) {
      const CellNodeId producer = *producer_it;
      const SpillPotential* candidate_potential = engine_.potential_spill_producers_.find(producer);
      if (candidate_potential == nullptr ||
          !SpillProducerIndex::admissible(workbook_, registry, producer, *candidate_potential, potential_arena)) {
        producer_it = attempted.erase(producer_it);
        continue;
      }
      if (closure.insert(producer).second) {
        bfs_queue.push_back(producer);
        expanded = true;
      }
      ++producer_it;
    }
    if (attempted.empty()) {
      attempted_it = attempted_candidates_.erase(attempted_it);
    } else {
      ++attempted_it;
    }
  }
  return expanded;
}

void RecalcEngine::DynamicReadPass::mark_stale_dirty() {
  for (const CellNodeId reader : stale_) {
    engine_.dirty_.mark(reader);
  }
}

bool RecalcEngine::drop_dynamic_only_cycles_locked(const std::vector<std::vector<CellNodeId>>& sccs) {
  if (!graph_.has_source_edges(DepGraph::DependencySource::kDynamicReference)) {
    return false;
  }
  bool dropped = false;
  for (const std::vector<CellNodeId>& component : sccs) {
    if (!is_cyclic_component(component, graph_) || !closes_through_dynamic_edge(component, graph_)) {
      continue;
    }
    for (const CellNodeId member : component) {
      graph_.clear_dynamic_dependencies_of(member);
    }
    dropped = true;
  }
  return dropped;
}

}  // namespace formulon::eval
