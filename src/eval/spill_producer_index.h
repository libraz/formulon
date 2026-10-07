//
// Index of formula anchors whose AST may produce an array, grouped by sheet,
// plus the geometric blocker probe used to filter them against a spill target.

#ifndef FORMULON_EVAL_SPILL_PRODUCER_INDEX_H_
#define FORMULON_EVAL_SPILL_PRODUCER_INDEX_H_

#include <cstddef>
#include <cstdint>
#include <unordered_map>
#include <vector>

#include "eval/dep_graph.h"
#include "eval/spill_potential.h"

namespace formulon {

class Arena;
struct Cell;
class Sheet;
class Workbook;

namespace eval {

class FunctionRegistry;

/// A stored formula or non-blank literal blocks every spill covering it. A phantom does not: its producer may shrink.
bool blocks_spill(const Cell* cell) noexcept;

/// A coordinate a not-yet-committed spill might fill, with the nearest stored blocker straight above and left of it.
struct SpillTarget {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::int64_t blocker_row = -1;
  std::int64_t blocker_col = -1;

  SpillTarget(const Sheet& sheet, std::uint32_t target_row, std::uint32_t target_col) noexcept;

  /// A spill covering the target covers the whole producer-to-target rectangle, so a blocker inside it rules the
  /// producer out unless the blocker is the producer's own cell.
  bool reachable_from(CellNodeId producer) const noexcept;
};

/// Formula anchors grouped by sheet whose AST may produce an array. This is the bounded candidate source for partial
/// range expansion; unlike the Sheet row map it never scans unrelated stored cells per wave. Not thread-safe; the
/// owning engine serialises access under its mutex.
class SpillProducerIndex {
 public:
  using SheetProducers = std::unordered_map<CellNodeId, SpillPotential, CellNodeIdHash>;

  /// Records `potential` for `cell`; `kNever` drops the entry.
  void update(CellNodeId cell, SpillPotential potential);

  void clear() noexcept { by_sheet_.clear(); }

  /// True when at least one producer was ever recorded at or past `sheet_id`'s slot.
  bool has_sheet(std::size_t sheet_id) const noexcept { return sheet_id < by_sheet_.size(); }

  /// Producers on one sheet; requires `has_sheet(sheet_id)`.
  const SheetProducers& on_sheet(std::size_t sheet_id) const { return by_sheet_[sheet_id]; }

  /// Recorded potential of `producer`, or nullptr when it is not indexed.
  const SpillPotential* find(CellNodeId producer) const;

  /// Re-classifies a candidate against the registry; `kNeedsRegistry` entries are re-parsed in `potential_arena`.
  static bool admissible(const Workbook& workbook, const FunctionRegistry& registry, CellNodeId producer,
                         SpillPotential static_potential, Arena& potential_arena);

 private:
  std::vector<SheetProducers> by_sheet_;
};

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SPILL_PRODUCER_INDEX_H_
