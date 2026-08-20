//
// Tracks the set of cells whose formula contains an Excel volatile
// function (NOW / TODAY / RAND / RANDBETWEEN / RANDARRAY / OFFSET /
// INDIRECT / INFO / CELL).
// On every recalc pass the engine seeds its `DirtySet` with this tracker
// so volatile-bearing cells are re-evaluated even when nothing they
// reference has changed.
//
// Registration also records *why* the cell is volatile, because the two
// reasons carry different scheduling constraints:
//
//   * value volatility (NOW / TODAY / RAND / RANDBETWEEN / RANDARRAY /
//     INFO / CELL) — the formula reads a fresh value out of the host
//     environment, but every cell it reads is described by a dependency
//     edge. The layering the parallel scheduler derives from those edges
//     is therefore complete, and the cell may be evaluated concurrently
//     with the rest of its layer like any ordinary formula.
//   * dynamic-reference volatility (OFFSET / INDIRECT) — the formula
//     computes its target at evaluation time, so no edge describes the
//     read and the target may sit in the same layer as the reader. Such a
//     cell must be evaluated against a quiescent cell store.
//
// Cells are registered after parsing: the recalc engine walks the AST,
// looks each call up via `is_volatile_function` /
// `is_dynamic_reference_function`, and calls `register_cell` if any call
// site matches. Re-parsing a cell first `unregister_cell`s the old entry
// then re-registers if the new AST is still volatile. Both classifiers
// match names case-insensitively over ASCII letters, so a hand-typed
// `=now()` is recognised without the caller having to canonicalize the
// lexeme first.

#ifndef FORMULON_EVAL_VOLATILE_TRACKER_H_
#define FORMULON_EVAL_VOLATILE_TRACKER_H_

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <unordered_map>
#include <utility>

#include "eval/dep_graph.h"

namespace formulon::eval {

/// Why a cell is volatile. `kDynamicReference` wins when a formula holds
/// both kinds of call, because the weaker guarantee is the binding one.
enum class VolatileKind : std::uint8_t {
  /// The result changes every pass, but the cells read are exactly the
  /// ones the dependency graph records.
  kValue,
  /// The formula resolves the cells it reads at evaluation time, so the
  /// dependency graph does not describe the read.
  kDynamicReference,
};

/// Set of cells known to invoke at least one volatile function, keyed by
/// the volatility class that decides how the cell may be scheduled.
class VolatileTracker {
 public:
  VolatileTracker() = default;

  // Move-only.
  VolatileTracker(const VolatileTracker&) = delete;
  VolatileTracker& operator=(const VolatileTracker&) = delete;
  VolatileTracker(VolatileTracker&&) noexcept = default;
  VolatileTracker& operator=(VolatileTracker&&) noexcept = default;
  ~VolatileTracker() = default;

  /// Marks `cell` as volatile with the supplied class. Idempotent for a
  /// given class; re-registering with a different one replaces the
  /// recorded class, matching how a re-parse rewrites the cell's entry.
  void register_cell(CellNodeId cell, VolatileKind kind) { cells_[cell] = kind; }

  /// Removes `cell` from the volatile set. No-op when absent.
  void unregister_cell(CellNodeId cell) { cells_.erase(cell); }

  /// Whether `cell` is currently registered as volatile, in either class.
  bool contains(CellNodeId cell) const { return cells_.find(cell) != cells_.end(); }

  /// Whether `cell` is registered as resolving its reads at evaluation
  /// time. False both for a value-volatile cell and for an unregistered
  /// one, so callers scheduling work can test this alone.
  bool contains_dynamic_reference(CellNodeId cell) const {
    const auto it = cells_.find(cell);
    return it != cells_.end() && it->second == VolatileKind::kDynamicReference;
  }

  /// Number of registered volatile cells, both classes together.
  std::size_t size() const noexcept { return cells_.size(); }

  /// Whether the set is empty.
  bool empty() const noexcept { return cells_.empty(); }

  /// Empties the set.
  void clear() noexcept { cells_.clear(); }

  /// Invokes `fn(CellNodeId)` for every registered cell, of either class.
  /// Iteration order is unspecified. `fn` must not mutate the tracker
  /// during iteration.
  template <typename Fn>
  void for_each(Fn&& visitor) const {
    for (const auto& entry : cells_) {
      std::forward<Fn>(visitor)(entry.first);
    }
  }

  /// Returns true iff `name` is one of the nine Excel volatile functions
  /// (`NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `RANDARRAY`, `OFFSET`,
  /// `INDIRECT`, `INFO`, `CELL`). Returns false for anything else,
  /// including the empty string.
  static bool is_volatile_function(std::string_view name);

  /// Returns true iff `name` is one of the volatile functions that
  /// resolves the cells it reads at evaluation time (`OFFSET`,
  /// `INDIRECT`). Every name accepted here is also accepted by
  /// `is_volatile_function`.
  static bool is_dynamic_reference_function(std::string_view name);

 private:
  std::unordered_map<CellNodeId, VolatileKind, CellNodeIdHash> cells_;
};

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_VOLATILE_TRACKER_H_
