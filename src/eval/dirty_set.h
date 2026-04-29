// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tiny set-of-cells helper used by the recalc engine to hold the dirty
// closure of a recalculation pass. The recalc loop seeds the set with the
// cells the user mutated (plus everything in `VolatileTracker`), then
// expands it by following `DepGraph::dependents_of` edges in BFS order
// until no new cells are added.
//
// The design is deliberately minimal: the dirty set is rebuilt from
// scratch each recalc pass (via `clear()`) and never persists across
// passes, so a plain `unordered_set` is sufficient. Adding atomic-friendly
// or sharded variants is reserved for the multi-threaded scheduler that
// lands later.

#ifndef FORMULON_EVAL_DIRTY_SET_H_
#define FORMULON_EVAL_DIRTY_SET_H_

#include <cstddef>
#include <unordered_set>
#include <utility>

#include "eval/dep_graph.h"

namespace formulon::eval {

/// Unordered set of `CellNodeId`s, used as a recalc-pass-scoped buffer of
/// cells that need to be re-evaluated. Not thread-safe.
class DirtySet {
 public:
  DirtySet() = default;

  // Move-only.
  DirtySet(const DirtySet&) = delete;
  DirtySet& operator=(const DirtySet&) = delete;
  DirtySet(DirtySet&&) noexcept = default;
  DirtySet& operator=(DirtySet&&) noexcept = default;
  ~DirtySet() = default;

  /// Inserts `cell` into the set. No-op when `cell` is already present.
  void mark(CellNodeId cell) { cells_.insert(cell); }

  /// Whether `cell` has been marked dirty.
  bool contains(CellNodeId cell) const { return cells_.find(cell) != cells_.end(); }

  /// Removes `cell` from the set; no-op when absent.
  void unmark(CellNodeId cell) { cells_.erase(cell); }

  /// Empties the set.
  void clear() noexcept { cells_.clear(); }

  /// Number of cells currently marked dirty.
  std::size_t size() const noexcept { return cells_.size(); }

  /// Whether the set is empty.
  bool empty() const noexcept { return cells_.empty(); }

  /// Invokes `fn(CellNodeId)` for every marked cell. Iteration order is
  /// unspecified. `fn` must not mutate the set during iteration.
  template <typename Fn>
  void for_each(Fn&& visitor) const {
    for (CellNodeId cell : cells_) {
      std::forward<Fn>(visitor)(cell);
    }
  }

 private:
  std::unordered_set<CellNodeId, CellNodeIdHash> cells_;
};

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_DIRTY_SET_H_
