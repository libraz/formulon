//
// Spill-release bookkeeping shared by the serial `RecalcEngine::recalc` and
// the parallel scheduler. Both drive the same release loop: a blocked spill
// anchor is queued while the workbook mutates, the loop wakes the queued
// producers, and a wave that leaves the blocked geometry and the woken target
// set unchanged is treated as no progress.
//
// The two entry points differ only in how they own the workbook lock, so the
// queue, the progress snapshot and the wave ceilings live here and are
// compiled once.

#ifndef FORMULON_EVAL_SPILL_RELEASE_H_
#define FORMULON_EVAL_SPILL_RELEASE_H_

#include <cstddef>
#include <cstdint>
#include <mutex>
#include <unordered_set>
#include <vector>

#include "eval/dep_graph.h"
#include "sheet.h"

namespace formulon {
class Workbook;
}  // namespace formulon

namespace formulon::eval::detail {

/// Collects the blocked spill anchors a mutation intersected, de-duplicated
/// and in first-seen order. `queue_spill_release` is invoked from inside the
/// workbook's own write path, so the queue carries its own mutex rather than
/// relying on the caller's.
struct SpillReleaseQueue {
  const Workbook* workbook = nullptr;
  std::mutex mutex;
  std::vector<CellNodeId> anchors;
  std::unordered_set<CellNodeId, CellNodeIdHash> queued;

  explicit SpillReleaseQueue(const Workbook* owner) : workbook(owner) {}

  std::vector<CellNodeId> take() {
    std::lock_guard<std::mutex> guard(mutex);
    std::vector<CellNodeId> out;
    out.swap(anchors);
    queued.clear();
    return out;
  }
};

/// `SpillReleaseCallback` implementation that records into the
/// `SpillReleaseQueue` passed as `raw`.
void queue_spill_release(void* raw, const Sheet& sheet, std::uint32_t first_row, std::uint32_t first_col,
                         std::uint32_t rows, std::uint32_t cols) noexcept;

/// One blocked spill footprint together with the sheet it sits on. Compared
/// field by field so the progress guard never depends on unordered-map
/// iteration order.
struct BlockedSpillState {
  std::uint16_t sheet_id = 0;
  BlockedSpillFootprint footprint;

  bool operator==(const BlockedSpillState& other) const noexcept {
    return sheet_id == other.sheet_id && footprint.anchor_row == other.footprint.anchor_row &&
           footprint.anchor_col == other.footprint.anchor_col && footprint.rows == other.footprint.rows &&
           footprint.cols == other.footprint.cols;
  }
};

/// Every blocked spill footprint in the workbook, in a deterministic order.
///
/// A release wave is allowed to change the blocked footprint set while it
/// wakes producers. If the same set comes back unchanged, however, the release
/// is not making progress — a common example is a volatile producer that
/// re-commits the same spill while another producer stays blocked.
std::vector<BlockedSpillState> snapshot_blocked_spills(const Workbook& workbook);

/// The woken anchors in a deterministic order, so parallel worker completion
/// order cannot manufacture a false difference between two waves.
///
/// The target set is part of progress: an unchanged blocked geometry may still
/// be advancing when a different producer is woken. Formula values are
/// intentionally not compared, because collision state is represented by the
/// spill / merge / occupied-cell geometry; literal edits update the release
/// targets through the workbook mutator.
std::vector<CellNodeId> canonical_release_targets(const std::vector<CellNodeId>& targets);

/// Ceilings that stop a release loop that cannot converge.
constexpr std::size_t kMaxSpillReleaseWaves = 4096U;
constexpr std::size_t kMaxNoProgressSpillWaves = 8U;

}  // namespace formulon::eval::detail

#endif  // FORMULON_EVAL_SPILL_RELEASE_H_
