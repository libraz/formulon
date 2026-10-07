//
// Private spill-table layout and the rectangle queries shared by the `Sheet`
// translation units. Internal to the cell store; not part of the public API.

#ifndef FORMULON_SHEET_SPILL_TABLE_H_
#define FORMULON_SHEET_SPILL_TABLE_H_

#include <algorithm>
#include <cstdint>
#include <unordered_map>
#include <vector>

#include "cell.h"
#include "sheet.h"

namespace formulon {

// ---------------------------------------------------------------------------
// Private spill-table layout
// ---------------------------------------------------------------------------
//
// The table is keyed only by anchor cell and owns each region's payload.
// Phantom lookup scans these rectangles rather than retaining one hash-map
// entry per spilled cell: a 1,000 x 100 spill is one region, not 99,999
// duplicate reverse-index nodes. A row-band index narrows that scan to the
// regions near the queried rows. Iteration order is undefined; consumers
// that need a deterministic order must sort externally.
struct SpillTable {
  // The band index points into `by_anchor`, so a copy would dangle.
  SpillTable() = default;
  SpillTable(const SpillTable&) = delete;
  SpillTable& operator=(const SpillTable&) = delete;

  // Mutate only through `add_region` / `erase_region` / `clear_regions`, which keep the band index in step.
  std::unordered_map<CellAddress, SpillRegion, CellAddressHash> by_anchor;
  // A failed dynamic-array commit leaves the anchor's attempted rectangle
  // here so later writes/removals inside that rectangle can re-dirty the
  // producer without requiring the user to touch the formula again.
  std::unordered_map<CellAddress, BlockedSpillFootprint, CellAddressHash> blocked_by_anchor;

  // Rows per band of the committed-region index. A region spanning more than
  // `kMaxIndexedBands` bands (a whole-column spill) goes to `tall_regions`,
  // which every query visits, so insertion stays bounded.
  static constexpr std::uint32_t kBandRows = 64U;
  static constexpr std::uint64_t kMaxIndexedBands = 64U;

  // Map nodes are stable across rehash, so the index holds region pointers.
  std::unordered_map<std::uint32_t, std::vector<const SpillRegion*>> regions_by_band;
  std::vector<const SpillRegion*> tall_regions;

  bool add_region(const CellAddress& anchor, SpillRegion region) {
    const auto inserted = by_anchor.emplace(anchor, std::move(region));
    if (!inserted.second) {
      return false;
    }
    const SpillRegion* stored = &inserted.first->second;
    const std::uint32_t first_band = first_band_of(*stored);
    const std::uint32_t last_band = last_band_of(*stored);
    if (static_cast<std::uint64_t>(last_band) - first_band + 1U > kMaxIndexedBands) {
      tall_regions.push_back(stored);
      return true;
    }
    for (std::uint32_t band = first_band; band <= last_band; ++band) {
      regions_by_band[band].push_back(stored);
    }
    return true;
  }

  void erase_region(std::unordered_map<CellAddress, SpillRegion, CellAddressHash>::iterator it) noexcept {
    const SpillRegion* stored = &it->second;
    const std::uint32_t first_band = first_band_of(*stored);
    const std::uint32_t last_band = last_band_of(*stored);
    if (static_cast<std::uint64_t>(last_band) - first_band + 1U > kMaxIndexedBands) {
      tall_regions.erase(std::remove(tall_regions.begin(), tall_regions.end(), stored), tall_regions.end());
    } else {
      for (std::uint32_t band = first_band; band <= last_band; ++band) {
        const auto bucket = regions_by_band.find(band);
        if (bucket == regions_by_band.end()) {
          continue;
        }
        auto& regions = bucket->second;
        regions.erase(std::remove(regions.begin(), regions.end(), stored), regions.end());
        if (regions.empty()) {
          regions_by_band.erase(bucket);
        }
      }
    }
    by_anchor.erase(it);
  }

  void clear_regions() noexcept {
    by_anchor.clear();
    regions_by_band.clear();
    tall_regions.clear();
  }

  // Calls `fn(const SpillRegion&)` for every committed region that may
  // intersect rows `[row_begin, row_end)`, possibly more than once, until it
  // returns true. Returns whether it did.
  template <typename Fn>
  bool any_region_near_rows(std::uint64_t row_begin, std::uint64_t row_end, Fn&& fn) const {
    for (const SpillRegion* region : tall_regions) {
      if (fn(*region)) {
        return true;
      }
    }
    if (row_end <= row_begin) {
      return false;
    }
    const auto first_band = static_cast<std::uint32_t>(row_begin / kBandRows);
    const auto last_band = static_cast<std::uint32_t>((row_end - 1U) / kBandRows);
    if (static_cast<std::uint64_t>(last_band) - first_band + 1U <= regions_by_band.size()) {
      for (std::uint32_t band = first_band; band <= last_band; ++band) {
        const auto bucket = regions_by_band.find(band);
        if (bucket == regions_by_band.end()) {
          continue;
        }
        for (const SpillRegion* region : bucket->second) {
          if (fn(*region)) {
            return true;
          }
        }
      }
      return false;
    }
    // A query spanning more bands than are populated walks the populated ones instead.
    for (const auto& [band, regions] : regions_by_band) {
      if (band < first_band || band > last_band) {
        continue;
      }
      for (const SpillRegion* region : regions) {
        if (fn(*region)) {
          return true;
        }
      }
    }
    return false;
  }

 private:
  static std::uint32_t first_band_of(const SpillRegion& region) noexcept { return region.anchor_row / kBandRows; }
  static std::uint32_t last_band_of(const SpillRegion& region) noexcept {
    return static_cast<std::uint32_t>((static_cast<std::uint64_t>(region.anchor_row) + region.rows - 1U) / kBandRows);
  }
};

/// True when two half-open rectangles share at least one coordinate.
///
/// Every rectangle in the cell store arrives as an origin plus an extent, and the
/// ends are widened to 64 bits so a rectangle touching the last row or column
/// cannot wrap. Written once because the spill table, the merge list, the
/// blocked-footprint table and the bulk read all need the same test, and a
/// hand-rolled copy per caller is a place for one comparison to drift.
inline bool RectsIntersect(std::uint64_t a_row_begin, std::uint64_t a_row_end, std::uint64_t a_col_begin,
                           std::uint64_t a_col_end, std::uint64_t b_row_begin, std::uint64_t b_row_end,
                           std::uint64_t b_col_begin, std::uint64_t b_col_end) noexcept {
  return a_row_begin < b_row_end && b_row_begin < a_row_end && a_col_begin < b_col_end && b_col_begin < a_col_end;
}

/// `RectsIntersect` for a spill-table rectangle, which is always an anchor
/// plus a `rows` x `cols` extent.
template <typename Rect>
bool RectIntersectsSpan(const Rect& rect, std::uint64_t row_begin, std::uint64_t row_end, std::uint64_t col_begin,
                        std::uint64_t col_end) noexcept {
  return RectsIntersect(row_begin, row_end, col_begin, col_end, rect.anchor_row,
                        static_cast<std::uint64_t>(rect.anchor_row) + rect.rows, rect.anchor_col,
                        static_cast<std::uint64_t>(rect.anchor_col) + rect.cols);
}

/// Returns the committed region whose rectangle contains `(row, col)`, anchor
/// included, or null. Callers hold the spill mutex.
inline const SpillRegion* find_region_covering(const SpillTable& table, std::uint32_t row, std::uint32_t col) {
  // Committed regions never overlap, so the first covering one is the only one.
  const SpillRegion* covering = nullptr;
  const std::uint64_t row_end = static_cast<std::uint64_t>(row) + 1U;
  table.any_region_near_rows(row, row_end, [&](const SpillRegion& region) {
    if (!RectIntersectsSpan(region, row, row_end, col, static_cast<std::uint64_t>(col) + 1U)) {
      return false;
    }
    covering = &region;
    return true;
  });
  return covering;
}

}  // namespace formulon

#endif  // FORMULON_SHEET_SPILL_TABLE_H_
