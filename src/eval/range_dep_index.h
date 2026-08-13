//
// Owner table for the compact rectangle dependencies the dep extractor
// emits (whole-column / whole-row references and every bounded rectangle
// above `kMaxMaterializedDependencyCells`).
//
// The recalc engine issues three queries against it:
//
//   * "which formulas watch a rectangle covering this cell?"
//     (`for_each_owner_covering`) — once per workbook cell write, so it must
//     not scan every registered rectangle;
//   * "which rectangles are watched, and by whom?"
//     (`for_each_distinct_range`) — once per spill reconcile and once per
//     partial-recalc closure expansion;
//   * "which rectangles does this formula watch?" (`for_each_range_of_owner`)
//     — trace expansion, and re-registration bookkeeping.
//
// Two structures keep those cheap. Rectangles are *interned*: a lookup
// dragged down a column authors one rectangle and N owner slots rather than
// N rectangles. Coverage lookups consult a coarse row-band bucket keyed on
// `(sheet_id, row / kBandRows)`, so a write inspects only the rectangles
// sharing its band. A bounded rectangle occupies `span / kBandRows + 1`
// buckets and a whole-column rectangle a fixed `kMaxRows / kBandRows`, which
// keeps insertion bounded regardless of the referenced area.
//
// Not thread-safe: the recalc engine owns the instance and mutates it under
// its own mutex.

#ifndef FORMULON_EVAL_RANGE_DEP_INDEX_H_
#define FORMULON_EVAL_RANGE_DEP_INDEX_H_

#include <cstddef>
#include <cstdint>
#include <unordered_map>
#include <utility>
#include <vector>

#include "eval/dep_extractor.h"
#include "eval/dep_graph.h"

namespace formulon::eval {

/// Many-to-many index between formula cells and the compact rectangles they
/// depend on.
class RangeDepIndex {
 public:
  /// Rows covered by one band bucket. Chosen so a whole-column rectangle
  /// touches 256 buckets — a fixed, small insertion cost — while an ordinary
  /// data-table rectangle of a few tens of thousands of rows lands in a
  /// handful.
  static constexpr std::uint32_t kBandRows = 4096U;

  RangeDepIndex() = default;

  // Move-only, matching the other recalc-engine-owned containers.
  RangeDepIndex(const RangeDepIndex&) = delete;
  RangeDepIndex& operator=(const RangeDepIndex&) = delete;
  RangeDepIndex(RangeDepIndex&&) noexcept = default;
  RangeDepIndex& operator=(RangeDepIndex&&) noexcept = default;
  ~RangeDepIndex() = default;

  /// Records that `owner` watches `range`. Idempotent: repeating the same
  /// pair leaves the index unchanged.
  void add(CellNodeId owner, const CellRangeDependency& range);

  /// Drops every rectangle watched by `owner`. Rectangles left without an
  /// owner are released. Safe on an owner that was never added.
  void erase_owner(CellNodeId owner);

  /// Empties the index.
  void clear() noexcept;

  /// Whether no owner is registered.
  bool empty() const noexcept { return ranges_by_owner_.empty(); }

  /// Number of distinct rectangles currently watched.
  std::size_t distinct_range_count() const noexcept { return live_range_count_; }

  /// Invokes `fn(CellNodeId owner)` for every formula whose rectangle covers
  /// `cell`. An owner watching two covering rectangles is visited twice; both
  /// call sites (dirty marking, edge insertion) are idempotent. `fn` must not
  /// mutate the index.
  template <typename Fn>
  void for_each_owner_covering(CellNodeId cell, Fn&& fn) const {
    const auto band = bands_.find(band_key(cell.sheet_id, cell.row));
    if (band == bands_.end()) {
      return;
    }
    for (const std::uint32_t range_id : band->second) {
      const RangeEntry& entry = ranges_[range_id];
      if (!entry.range.contains(cell)) {
        continue;
      }
      for (const CellNodeId owner : entry.owners) {
        fn(owner);
      }
    }
  }

  /// Invokes `fn(std::uint32_t range_id, const CellRangeDependency& range,
  /// const std::vector<CellNodeId>& owners)` once per distinct rectangle.
  /// `range_id` is stable until the index is mutated, so callers running a
  /// fixed point can use it as a "already processed" key. `fn` must not
  /// mutate the index.
  template <typename Fn>
  void for_each_distinct_range(Fn&& fn) const {
    for (std::uint32_t range_id = 0; range_id < ranges_.size(); ++range_id) {
      const RangeEntry& entry = ranges_[range_id];
      if (entry.owners.empty()) {
        continue;  // Released slot awaiting reuse.
      }
      fn(range_id, entry.range, entry.owners);
    }
  }

  /// Invokes `fn(const CellRangeDependency&)` for every rectangle `owner`
  /// watches. `fn` must not mutate the index.
  template <typename Fn>
  void for_each_range_of_owner(CellNodeId owner, Fn&& fn) const {
    const auto it = ranges_by_owner_.find(owner);
    if (it == ranges_by_owner_.end()) {
      return;
    }
    for (const OwnedRange& owned : it->second) {
      fn(ranges_[owned.range_id].range);
    }
  }

 private:
  // One interned rectangle. An entry whose owner list is empty is a released
  // slot: its id lives in `free_range_ids_` and is skipped by iteration.
  struct RangeEntry {
    CellRangeDependency range;
    std::vector<CellNodeId> owners;
  };

  // Back-reference from an owner to its slot inside `RangeEntry::owners`, so
  // removal is O(1) instead of a scan over a rectangle watched by every
  // formula in a dragged column.
  struct OwnedRange {
    std::uint32_t range_id = 0;
    std::uint32_t owner_index = 0;
  };

  static std::uint64_t band_key(std::uint16_t sheet_id, std::uint32_t row) noexcept {
    return (static_cast<std::uint64_t>(sheet_id) << 32U) | (row / kBandRows);
  }

  std::uint32_t intern_range(const CellRangeDependency& range);
  void release_range(std::uint32_t range_id);

  std::vector<RangeEntry> ranges_;
  std::vector<std::uint32_t> free_range_ids_;
  std::unordered_map<CellRangeDependency, std::uint32_t, CellRangeDependencyHash> range_ids_;
  std::unordered_map<CellNodeId, std::vector<OwnedRange>, CellNodeIdHash> ranges_by_owner_;
  std::unordered_map<std::uint64_t, std::vector<std::uint32_t>> bands_;
  std::size_t live_range_count_ = 0U;
};

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_RANGE_DEP_INDEX_H_
