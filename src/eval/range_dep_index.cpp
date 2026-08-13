//
// Implementation of `RangeDepIndex`. See `range_dep_index.h` for the
// contract.

#include "eval/range_dep_index.h"

#include <algorithm>
#include <cstdint>
#include <utility>
#include <vector>

namespace formulon::eval {

std::uint32_t RangeDepIndex::intern_range(const CellRangeDependency& range) {
  const auto existing = range_ids_.find(range);
  if (existing != range_ids_.end()) {
    return existing->second;
  }

  std::uint32_t range_id = 0;
  if (!free_range_ids_.empty()) {
    range_id = free_range_ids_.back();
    free_range_ids_.pop_back();
    ranges_[range_id].range = range;
  } else {
    range_id = static_cast<std::uint32_t>(ranges_.size());
    ranges_.push_back(RangeEntry{range, {}});
  }
  range_ids_.emplace(range, range_id);
  const std::uint32_t first_band = range.row_first / kBandRows;
  const std::uint32_t last_band = range.row_last / kBandRows;
  for (std::uint32_t band = first_band; band <= last_band; ++band) {
    bands_[(static_cast<std::uint64_t>(range.sheet_id) << 32U) | band].push_back(range_id);
  }
  ++live_range_count_;
  return range_id;
}

void RangeDepIndex::release_range(std::uint32_t range_id) {
  RangeEntry& entry = ranges_[range_id];
  const std::uint32_t first_band = entry.range.row_first / kBandRows;
  const std::uint32_t last_band = entry.range.row_last / kBandRows;
  for (std::uint32_t band = first_band; band <= last_band; ++band) {
    const std::uint64_t key = (static_cast<std::uint64_t>(entry.range.sheet_id) << 32U) | band;
    const auto bucket = bands_.find(key);
    if (bucket == bands_.end()) {
      continue;
    }
    auto& ids = bucket->second;
    ids.erase(std::remove(ids.begin(), ids.end(), range_id), ids.end());
    if (ids.empty()) {
      bands_.erase(bucket);
    }
  }
  range_ids_.erase(entry.range);
  entry.range = CellRangeDependency{};
  free_range_ids_.push_back(range_id);
  --live_range_count_;
}

void RangeDepIndex::add(CellNodeId owner, const CellRangeDependency& range) {
  const std::uint32_t range_id = intern_range(range);
  std::vector<OwnedRange>& owned = ranges_by_owner_[owner];
  for (const OwnedRange& existing : owned) {
    if (existing.range_id == range_id) {
      return;  // Already watching this rectangle.
    }
  }
  std::vector<CellNodeId>& owners = ranges_[range_id].owners;
  owned.push_back(OwnedRange{range_id, static_cast<std::uint32_t>(owners.size())});
  owners.push_back(owner);
}

void RangeDepIndex::erase_owner(CellNodeId owner) {
  const auto it = ranges_by_owner_.find(owner);
  if (it == ranges_by_owner_.end()) {
    return;
  }

  for (const OwnedRange& owned : it->second) {
    std::vector<CellNodeId>& owners = ranges_[owned.range_id].owners;
    const std::size_t last_index = owners.size() - 1U;
    if (owned.owner_index != last_index) {
      // Swap-remove: the owner moved into the vacated slot must learn its new
      // index, otherwise its own erase would target a stranger's slot.
      const CellNodeId moved = owners[last_index];
      owners[owned.owner_index] = moved;
      // `find`, not `operator[]`: inserting here could rehash the map and
      // invalidate the owner list this loop is walking. The moved owner is
      // registered by construction, so a miss is impossible.
      const auto moved_it = ranges_by_owner_.find(moved);
      if (moved_it != ranges_by_owner_.end()) {
        for (OwnedRange& moved_owned : moved_it->second) {
          if (moved_owned.range_id == owned.range_id) {
            moved_owned.owner_index = owned.owner_index;
            break;
          }
        }
      }
    }
    owners.pop_back();
    if (owners.empty()) {
      release_range(owned.range_id);
    }
  }
  ranges_by_owner_.erase(it);
}

void RangeDepIndex::clear() noexcept {
  ranges_.clear();
  free_range_ids_.clear();
  range_ids_.clear();
  ranges_by_owner_.clear();
  bands_.clear();
  live_range_count_ = 0U;
}

}  // namespace formulon::eval
