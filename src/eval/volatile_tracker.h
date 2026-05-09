// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tracks the set of cells whose formula contains an Excel volatile
// function (NOW / TODAY / RAND / RANDBETWEEN / RANDARRAY / OFFSET /
// INDIRECT / INFO / CELL).
// On every recalc pass the engine seeds its `DirtySet` with this tracker
// so volatile-bearing cells are re-evaluated even when nothing they
// reference has changed.
//
// Cells are registered after parsing: the recalc engine walks the AST,
// looks each call up via `is_volatile_function`, and calls
// `register_cell` if any call site matches. Re-parsing a cell first
// `unregister_cell`s the old entry then re-registers if the new AST is
// still volatile. `is_volatile_function` is a pure ASCII-uppercase string
// match (matching how the function registry is keyed); callers that have
// a mixed-case name must canonicalize it first.

#ifndef FORMULON_EVAL_VOLATILE_TRACKER_H_
#define FORMULON_EVAL_VOLATILE_TRACKER_H_

#include <cstddef>
#include <string_view>
#include <unordered_set>
#include <utility>

#include "eval/dep_graph.h"

namespace formulon::eval {

/// Set of cells known to invoke at least one volatile function.
///
/// The tracker stores no metadata about *which* volatile function caused
/// the registration — that is irrelevant to the recalc engine, which only
/// needs to know "is this cell always-dirty at the start of a pass?".
class VolatileTracker {
 public:
  VolatileTracker() = default;

  // Move-only.
  VolatileTracker(const VolatileTracker&) = delete;
  VolatileTracker& operator=(const VolatileTracker&) = delete;
  VolatileTracker(VolatileTracker&&) noexcept = default;
  VolatileTracker& operator=(VolatileTracker&&) noexcept = default;
  ~VolatileTracker() = default;

  /// Marks `cell` as volatile. Idempotent.
  void register_cell(CellNodeId cell) { cells_.insert(cell); }

  /// Removes `cell` from the volatile set. No-op when absent.
  void unregister_cell(CellNodeId cell) { cells_.erase(cell); }

  /// Whether `cell` is currently registered as volatile.
  bool contains(CellNodeId cell) const { return cells_.find(cell) != cells_.end(); }

  /// Number of registered volatile cells.
  std::size_t size() const noexcept { return cells_.size(); }

  /// Whether the set is empty.
  bool empty() const noexcept { return cells_.empty(); }

  /// Empties the set.
  void clear() noexcept { cells_.clear(); }

  /// Invokes `fn(CellNodeId)` for every registered cell. Iteration order
  /// is unspecified. `fn` must not mutate the tracker during iteration.
  template <typename Fn>
  void for_each(Fn&& visitor) const {
    for (CellNodeId cell : cells_) {
      std::forward<Fn>(visitor)(cell);
    }
  }

  /// Returns true iff `name` is one of the nine Excel volatile functions
  /// (`NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `RANDARRAY`, `OFFSET`,
  /// `INDIRECT`, `INFO`, `CELL`). Match is case-sensitive and assumes
  /// `name` is already canonicalized to ASCII uppercase, mirroring how
  /// names are keyed in `FunctionRegistry`. Returns false for anything
  /// else — including lowercase aliases and the empty string.
  static bool is_volatile_function(std::string_view name);

 private:
  std::unordered_set<CellNodeId, CellNodeIdHash> cells_;
};

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_VOLATILE_TRACKER_H_
