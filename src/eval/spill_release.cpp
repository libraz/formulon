//
// Implementation of the shared spill-release bookkeeping. See
// `spill_release.h` for the contract.

#include "eval/spill_release.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <mutex>
#include <vector>

#include "cell.h"
#include "eval/dep_graph.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon::eval::detail {

void queue_spill_release(void* raw, const Sheet& sheet, std::uint32_t first_row, std::uint32_t first_col,
                         std::uint32_t rows, std::uint32_t cols) noexcept {
  auto* queue = static_cast<SpillReleaseQueue*>(raw);
  if (queue == nullptr || queue->workbook == nullptr) {
    return;
  }
  std::size_t sheet_index = queue->workbook->sheet_count();
  for (std::size_t i = 0; i < queue->workbook->sheet_count(); ++i) {
    if (&queue->workbook->sheet(i) == &sheet) {
      sheet_index = i;
      break;
    }
  }
  if (sheet_index >= queue->workbook->sheet_count()) {
    return;
  }
  const std::vector<CellAddress> anchors = sheet.blocked_spill_anchors_intersecting(first_row, first_col, rows, cols);
  std::lock_guard<std::mutex> guard(queue->mutex);
  for (const CellAddress anchor : anchors) {
    const CellNodeId node{static_cast<std::uint16_t>(sheet_index), anchor.row, anchor.col};
    if (queue->queued.insert(node).second) {
      queue->anchors.push_back(node);
    }
  }
}

std::vector<BlockedSpillState> snapshot_blocked_spills(const Workbook& workbook) {
  std::vector<BlockedSpillState> state;
  for (std::size_t sheet_id = 0; sheet_id < workbook.sheet_count(); ++sheet_id) {
    for (const BlockedSpillFootprint& footprint : workbook.sheet(sheet_id).blocked_spill_footprints()) {
      state.push_back(BlockedSpillState{static_cast<std::uint16_t>(sheet_id), footprint});
    }
  }
  std::sort(state.begin(), state.end(), [](const BlockedSpillState& lhs, const BlockedSpillState& rhs) {
    if (lhs.sheet_id != rhs.sheet_id) {
      return lhs.sheet_id < rhs.sheet_id;
    }
    if (lhs.footprint.anchor_row != rhs.footprint.anchor_row) {
      return lhs.footprint.anchor_row < rhs.footprint.anchor_row;
    }
    if (lhs.footprint.anchor_col != rhs.footprint.anchor_col) {
      return lhs.footprint.anchor_col < rhs.footprint.anchor_col;
    }
    if (lhs.footprint.rows != rhs.footprint.rows) {
      return lhs.footprint.rows < rhs.footprint.rows;
    }
    return lhs.footprint.cols < rhs.footprint.cols;
  });
  return state;
}

std::vector<CellNodeId> canonical_release_targets(const std::vector<CellNodeId>& targets) {
  std::vector<CellNodeId> sorted = targets;
  std::sort(sorted.begin(), sorted.end(), CellNodeIdOrder{});
  return sorted;
}

}  // namespace formulon::eval::detail
