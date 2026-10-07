//
// Out-of-line implementation of `Sheet`'s dynamic-array spill table: anchor
// and footprint queries, collision probes, commit, and clearing.

#include <algorithm>
#include <cassert>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include "cell.h"
#include "sheet.h"
#include "sheet_spill_table.h"
#include "utils/resource_budget.h"
#include "value.h"

namespace formulon {

namespace {

/// Returns the anchors of every rectangle in `table` that overlaps
/// `[first_row, first_row + rows) x [first_col, first_col + cols)`.
///
/// Both anchor tables carry the same four rectangle fields under different
/// payload types, so the walk is shared and only the map type varies. `table`
/// may be null when the sheet has no spill table yet; callers hold the spill
/// mutex.
template <typename AnchorMap>
std::vector<CellAddress> collect_anchors_intersecting(const AnchorMap* table, std::uint32_t first_row,
                                                      std::uint32_t first_col, std::uint32_t rows, std::uint32_t cols) {
  std::vector<CellAddress> out;
  if (table == nullptr || rows == 0U || cols == 0U) {
    return out;
  }
  const std::uint64_t row_end = static_cast<std::uint64_t>(first_row) + rows;
  const std::uint64_t col_end = static_cast<std::uint64_t>(first_col) + cols;
  out.reserve(table->size());
  for (const auto& [address, rect] : *table) {
    if (RectIntersectsSpan(rect, first_row, row_end, first_col, col_end)) {
      out.push_back(address);
    }
  }
  return out;
}

// Returns true when `c` is "occupied" for the purposes of a spill collision
// check: a non-default cell (literal value or formula) lives there. The
// anchor cell of the would-be spill is excluded from this check by the
// caller. The caller resolves the cell pointer via `cell_at_locked` (while
// holding `spill_mutex_`) and passes it in, so this helper does not reach
// back into `Sheet` and cannot self-deadlock.
bool IsCellOccupied(const Cell* c) noexcept {
  if (c == nullptr) {
    return false;
  }
  if (!c->formula_text.empty()) {
    return true;
  }
  return !c->cached_value.is_blank();
}

// Deep-copies `cells` into `region`, interning every Text payload's bytes
// into `region.owned_strings` and rewriting the corresponding `Value` so its
// `string_view` points at the interned copy. Non-text cells are copied
// verbatim. The strings are reserved up-front so no later push_back can
// invalidate the string_view payloads of earlier cells: a string move from
// SSO to heap (or a vector reallocation) would otherwise corrupt every
// previously interned reference. The exact reservation is the count of Text
// cells in the input.
//
// Pass-by-const-ref is intentional: the input is conceptually consumed (the
// caller has just received it by value from `commit_spill`), but each `Value`
// is trivially copyable and the Text payload must be deep-copied byte-by-byte
// anyway, so a `std::move` of the outer vector would not save any work.
void CopyCellsWithOwnedText(const std::vector<Value>& src, SpillRegion& region) {
  std::size_t text_count = 0;
  for (const Value& v : src) {
    if (v.is_text()) {
      ++text_count;
    }
  }
  region.owned_strings.reserve(text_count);
  // `Value` has no public default constructor, so build the cells vector
  // by reservation + push_back rather than by `resize`.
  region.cells.reserve(src.size());
  for (const Value& v : src) {
    if (v.is_text()) {
      region.owned_strings.emplace_back(v.as_text());
      region.cells.push_back(Value::text(region.owned_strings.back()));
    } else {
      region.cells.push_back(v);
    }
  }
}

}  // namespace

void Sheet::for_each_spill_phantom(void (*visit)(CellAddress address, void* ctx), void* ctx) const {
  if (visit == nullptr) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (spill_table_ == nullptr) {
    return;
  }
  for (const auto& kv : spill_table_->by_anchor) {
    const SpillRegion& region = kv.second;
    for (std::uint32_t r = 0; r < region.rows; ++r) {
      for (std::uint32_t c = 0; c < region.cols; ++c) {
        if (r == 0U && c == 0U) {
          continue;  // Exclude the anchor; it has a real slot in `rows_`.
        }
        visit(CellAddress{region.anchor_row + r, region.anchor_col + c}, ctx);
      }
    }
  }
}

std::vector<CellAddress> Sheet::spill_phantom_addresses() const {
  std::vector<CellAddress> out;
  for_each_spill_phantom(
      [](CellAddress address, void* ctx) { static_cast<std::vector<CellAddress>*>(ctx)->push_back(address); }, &out);
  return out;
}

std::vector<CellAddress> Sheet::blocked_spill_anchors_intersecting(std::uint32_t first_row, std::uint32_t first_col,
                                                                   std::uint32_t rows, std::uint32_t cols) const {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  return collect_anchors_intersecting(spill_table_ == nullptr ? nullptr : &spill_table_->blocked_by_anchor, first_row,
                                      first_col, rows, cols);
}

std::vector<CellAddress> Sheet::committed_spill_anchors_intersecting(std::uint32_t first_row, std::uint32_t first_col,
                                                                     std::uint32_t rows, std::uint32_t cols) const {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  return collect_anchors_intersecting(spill_table_ == nullptr ? nullptr : &spill_table_->by_anchor, first_row,
                                      first_col, rows, cols);
}

std::vector<CellAddress> Sheet::blocked_spill_anchors() const {
  std::vector<CellAddress> out;
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (spill_table_ == nullptr) {
    return out;
  }
  out.reserve(spill_table_->blocked_by_anchor.size());
  for (const auto& [address, unused] : spill_table_->blocked_by_anchor) {
    (void)unused;
    out.push_back(address);
  }
  return out;
}

std::vector<BlockedSpillFootprint> Sheet::blocked_spill_footprints() const {
  std::vector<BlockedSpillFootprint> out;
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (spill_table_ == nullptr) {
    return out;
  }
  out.reserve(spill_table_->blocked_by_anchor.size());
  for (const auto& [unused, footprint] : spill_table_->blocked_by_anchor) {
    (void)unused;
    out.push_back(footprint);
  }
  return out;
}

std::optional<BlockedSpillFootprint> Sheet::committed_spill_footprint_covering(std::uint32_t row,
                                                                               std::uint32_t col) const {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (spill_table_ == nullptr) {
    return std::nullopt;
  }
  const SpillRegion* covering = find_region_covering(*spill_table_, row, col);
  if (covering == nullptr) {
    return std::nullopt;
  }
  return BlockedSpillFootprint{covering->anchor_row, covering->anchor_col, covering->rows, covering->cols};
}

std::vector<SpillFootprint> Sheet::committed_spill_footprints() const {
  std::vector<SpillFootprint> out;
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (spill_table_ == nullptr) {
    return out;
  }
  out.reserve(spill_table_->by_anchor.size());
  for (const auto& [unused, region] : spill_table_->by_anchor) {
    (void)unused;
    out.push_back(SpillFootprint{region.anchor_row, region.anchor_col, region.rows, region.cols});
  }
  return out;
}

void Sheet::restore_blocked_spill_footprints(std::vector<BlockedSpillFootprint> footprints) {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  for (const BlockedSpillFootprint& footprint : footprints) {
    if (footprint.rows == 0U || footprint.cols == 0U || !coord_in_grid(footprint.anchor_row, footprint.anchor_col) ||
        static_cast<std::uint64_t>(footprint.anchor_row) + footprint.rows > kMaxRows ||
        static_cast<std::uint64_t>(footprint.anchor_col) + footprint.cols > kMaxCols) {
      continue;
    }
    if (spill_table_ == nullptr) {
      spill_table_ = std::make_unique<SpillTable>();
    }
    spill_table_->blocked_by_anchor[CellAddress{footprint.anchor_row, footprint.anchor_col}] = footprint;
  }
}

// ---------------------------------------------------------------------------
// Spill API
// ---------------------------------------------------------------------------

const SpillRegion* Sheet::spill_region_at_anchor(std::uint32_t row, std::uint32_t col) const noexcept {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  // Not called from within any other locked Sheet method, so the body stays
  // inline rather than delegating to a `_locked` variant.
  if (spill_table_ == nullptr) {
    return nullptr;
  }
  const auto it = spill_table_->by_anchor.find(CellAddress{row, col});
  if (it == spill_table_->by_anchor.end()) {
    return nullptr;
  }
  return &it->second;
}

const SpillRegion* Sheet::spill_region_covering(std::uint32_t row, std::uint32_t col) const noexcept {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  return spill_region_covering_locked(row, col);
}

const SpillRegion* Sheet::spill_region_covering_locked(std::uint32_t row, std::uint32_t col) const noexcept {
  if (spill_table_ == nullptr) {
    return nullptr;
  }
  const SpillRegion* covering = find_region_covering(*spill_table_, row, col);
  if (covering != nullptr && row == covering->anchor_row && col == covering->anchor_col) {
    return nullptr;
  }
  return covering;
}

bool Sheet::spill_would_collide(std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint32_t rows,
                                std::uint32_t cols) const noexcept {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  return probe_spill_footprint_locked(anchor_row, anchor_col, rows, cols, /*scan_steps=*/nullptr) !=
         SpillAdmission::kAdmissible;
}

bool Sheet::footprint_holds_occupied_cell_locked(std::uint32_t anchor_row, std::uint32_t anchor_col,
                                                 std::uint64_t row_end, std::uint64_t col_end,
                                                 std::uint64_t* scan_steps) const noexcept {
  const auto count_step = [&]() noexcept {
    if (scan_steps != nullptr) {
      ++*scan_steps;
    }
  };

  // Inspect one stored row's materialised run where it overlaps the column
  // span. `RowCells` holds a dense run starting at `first_col()`, so the
  // overlap is a contiguous slice and the leading gap costs nothing.
  const auto row_holds_blocker = [&](std::uint32_t row, const RowCells& cells) noexcept {
    if (cells.empty()) {
      return false;
    }
    const std::size_t first = std::max<std::size_t>(anchor_col, cells.first_col());
    const std::size_t last = std::min<std::size_t>(static_cast<std::size_t>(col_end), cells.size());
    for (std::size_t col = first; col < last; ++col) {
      if (row == anchor_row && col == anchor_col) {
        continue;
      }
      count_step();
      if (IsCellOccupied(cells.find(static_cast<std::uint32_t>(col)))) {
        return true;
      }
    }
    return false;
  };

  // Whichever of "walk the stored rows" and "probe each row of the rectangle"
  // is smaller wins: the first is proportional to the sheet, the second to
  // the rectangle. A whole-column footprint spans every row of the grid, so
  // only the first strategy keeps it affordable. Mirrors the same choice in
  // `populated_extent`.
  const std::uint64_t rect_rows = row_end - anchor_row;
  if (static_cast<std::uint64_t>(rows_.size()) <= rect_rows) {
    for (const auto& [row, cells] : rows_) {
      count_step();
      if (static_cast<std::uint64_t>(row) < anchor_row || static_cast<std::uint64_t>(row) >= row_end) {
        continue;
      }
      if (row_holds_blocker(row, cells)) {
        return true;
      }
    }
    return false;
  }
  for (std::uint64_t row = anchor_row; row < row_end; ++row) {
    count_step();
    const auto it = rows_.find(static_cast<std::uint32_t>(row));
    if (it == rows_.end()) {
      continue;
    }
    if (row_holds_blocker(static_cast<std::uint32_t>(row), it->second)) {
      return true;
    }
  }
  return false;
}

Sheet::SpillAdmission Sheet::probe_spill_footprint(std::uint32_t anchor_row, std::uint32_t anchor_col,
                                                   std::uint32_t rows, std::uint32_t cols,
                                                   std::uint64_t* scan_steps) const noexcept {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  return probe_spill_footprint_locked(anchor_row, anchor_col, rows, cols, scan_steps);
}

Sheet::SpillAdmission Sheet::probe_spill_footprint_locked(std::uint32_t anchor_row, std::uint32_t anchor_col,
                                                          std::uint32_t rows, std::uint32_t cols,
                                                          std::uint64_t* scan_steps) const noexcept {
  if (scan_steps != nullptr) {
    *scan_steps = 0;
  }
  // Same geometry rule the collision predicate applies, reported as its own
  // verdict: a degenerate shape, an anchor off the grid, or a rectangle whose
  // far edge leaves the grid can never commit.
  if (rows == 0U || cols == 0U || !coord_in_grid(anchor_row, anchor_col)) {
    return SpillAdmission::kOutsideGrid;
  }
  const std::uint64_t row_end = static_cast<std::uint64_t>(anchor_row) + rows;
  const std::uint64_t col_end = static_cast<std::uint64_t>(anchor_col) + cols;
  if (row_end > kMaxRows || col_end > kMaxCols) {
    return SpillAdmission::kOutsideGrid;
  }

  // Spill rectangles and merged ranges are rectangle-intersection tests over
  // their own tables (spills through the row-band index), never the
  // footprint's area; only the stored-cell sweep needed the sparse walk.
  //
  // A pre-existing region at this anchor is the producer's own spill. It is
  // ignored wholesale: ad-hoc evaluation is read-only and cannot clear it,
  // while commit_spill clears it before reaching this predicate.
  if (spill_table_ != nullptr &&
      spill_table_->any_region_near_rows(anchor_row, row_end, [&](const SpillRegion& region) {
        return !(region.anchor_row == anchor_row && region.anchor_col == anchor_col) &&
               RectIntersectsSpan(region, anchor_row, row_end, anchor_col, col_end);
      })) {
    return SpillAdmission::kBlocked;
  }
  // Merged cells occupy their complete rectangle even when only the top-left
  // coordinate has a stored Cell. Any intersection is a blocker, including a
  // merge whose top-left cell is the requested spill anchor.
  for (const MergeRange& merge : merges_) {
    if (merge.first_row > merge.last_row || merge.first_col > merge.last_col) {
      continue;
    }
    if (RectsIntersect(anchor_row, row_end, anchor_col, col_end, merge.first_row,
                       static_cast<std::uint64_t>(merge.last_row) + 1U, merge.first_col,
                       static_cast<std::uint64_t>(merge.last_col) + 1U)) {
      return SpillAdmission::kBlocked;
    }
  }
  if (footprint_holds_occupied_cell_locked(anchor_row, anchor_col, row_end, col_end, scan_steps)) {
    return SpillAdmission::kBlocked;
  }
  return SpillAdmission::kAdmissible;
}

void Sheet::reject_spill_footprint(std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint32_t rows,
                                   std::uint32_t cols) {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  reject_spill_footprint_locked(anchor_row, anchor_col, rows, cols);
}

void Sheet::reject_spill_footprint_locked(std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint32_t rows,
                                          std::uint32_t cols) {
  if (!coord_in_grid(anchor_row, anchor_col)) {
    return;
  }
  // Surface #SPILL! at the anchor; preserve the existing literal and all
  // other metadata at the colliding cell.
  RowCells& row_cells = rows_[anchor_row];
  Cell& anchor_slot = row_cells.ensure(anchor_col);
  anchor_slot.cached_value = Value::error(ErrorCode::Spill);
  if (spill_table_ == nullptr) {
    spill_table_ = std::make_unique<SpillTable>();
  }
  // Remembering the rectangle is what lets the release machinery retry this
  // anchor once the blocker goes away; dropping it strands the #SPILL!.
  spill_table_->blocked_by_anchor[CellAddress{anchor_row, anchor_col}] =
      BlockedSpillFootprint{anchor_row, anchor_col, rows, cols};
  cell_enumeration_revision_.bump();
}

bool Sheet::commit_spill(std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint32_t rows, std::uint32_t cols,
                         std::vector<Value> cells) {
  // Shape validation. These conditions indicate caller bugs; report them
  // via the debug assert and refuse the registration so release builds
  // remain memory-safe.
  if (rows == 0U || cols == 0U) {
    assert(false && "commit_spill: zero-sized spill region");
    return false;
  }
  // Area first, in 64-bit, and before the payload length is derived from it:
  // this is the same ceiling the evaluator's array allocator applies to the
  // result it would spill, and every producer that reaches here has already
  // passed it, so it rejects nothing the engine can build. It exists so the
  // sheet does not depend on its caller having checked — a region this large
  // costs a `Value` per cell here plus one enumerated coordinate per phantom
  // in the C ABI, which is where a 32-bit host runs out of address space.
  // Establishing it first also keeps `rows * cols` from wrapping `size_t` on
  // such a host, which would let a short payload match a huge shape.
  if (static_cast<std::uint64_t>(rows) * cols > kMaxRangeExpansionCells) {
    assert(false && "commit_spill: footprint exceeds the dynamic-array cell ceiling");
    return false;
  }
  const std::size_t expected_size = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  if (cells.size() != expected_size) {
    assert(false && "commit_spill: cells.size() does not match rows*cols");
    return false;
  }
  if (anchor_row >= kMaxRows || anchor_col >= kMaxCols) {
    assert(false && "commit_spill: anchor out of bounds");
    return false;
  }
  if (static_cast<std::uint64_t>(anchor_row) + rows > kMaxRows ||
      static_cast<std::uint64_t>(anchor_col) + cols > kMaxCols) {
    assert(false && "commit_spill: footprint exceeds sheet bounds");
    return false;
  }

  // Hold `spill_mutex_` for the whole commit: the collision scan reads and
  // the registration writes `spill_table_` + `rows_` as one atomic unit so
  // two parallel-recalc workers spilling on the same sheet cannot interleave.
  // `std::mutex` is non-recursive, so every internal call below routes
  // through the `_locked` helpers rather than the public, self-locking ones.
  const std::lock_guard<std::mutex> guard(*spill_mutex_);

  // Drop any region currently anchored at this cell first, regardless of
  // whether the new commit ends up succeeding. The "register over an
  // existing region" case is intentionally idempotent.
  clear_spill_locked(anchor_row, anchor_col);

  // Admission is shared with the read-only/ad-hoc evaluation path and with
  // producers that probe before materialising, so a refusal here and a
  // refusal there record the same thing.
  if (probe_spill_footprint_locked(anchor_row, anchor_col, rows, cols, /*scan_steps=*/nullptr) !=
      SpillAdmission::kAdmissible) {
    reject_spill_footprint_locked(anchor_row, anchor_col, rows, cols);
    return false;
  }

  // Materialise the spill table on first use.
  if (spill_table_ == nullptr) {
    spill_table_ = std::make_unique<SpillTable>();
  }

  // Build the region with deep-copied Text payloads.
  SpillRegion region;
  region.anchor_row = anchor_row;
  region.anchor_col = anchor_col;
  region.rows = rows;
  region.cols = cols;
  CopyCellsWithOwnedText(cells, region);

  // Register the anchor entry. Capture the first cell up-front because the
  // region is about to be moved into the map; afterwards the by-value
  // `cells[0]` is no longer reachable through `region`.
  const Value first_cell = region.cells[0];
  const CellAddress anchor_addr{anchor_row, anchor_col};
  const bool inserted = spill_table_->add_region(anchor_addr, std::move(region));
  assert(inserted && "commit_spill: anchor entry already present after clear");
  (void)inserted;

  // Anchor's cached_value mirrors the first cell of the region so that
  // `cell_at(anchor)->cached_value` and `resolve_cell_value(anchor)` agree
  // without a special anchor case.
  RowCells& row_cells = rows_[anchor_row];
  Cell& anchor_slot = row_cells.ensure(anchor_col);
  anchor_slot.cached_value = first_cell;
  cell_enumeration_revision_.bump();
  return true;
}

void Sheet::shift_blocked_spills_locked(const StructuralEdit& edit) {
  if (spill_table_ == nullptr || spill_table_->blocked_by_anchor.empty()) {
    return;
  }
  std::unordered_map<CellAddress, BlockedSpillFootprint, CellAddressHash> shifted;
  shifted.reserve(spill_table_->blocked_by_anchor.size());
  const std::uint32_t bound = edit.row_axis ? kMaxRows : kMaxCols;
  const std::uint64_t delete_end = static_cast<std::uint64_t>(edit.index) + edit.count;
  for (const auto& [address, footprint] : spill_table_->blocked_by_anchor) {
    std::uint32_t coordinate = edit.row_axis ? address.row : address.col;
    if (edit.is_delete) {
      if (static_cast<std::uint64_t>(coordinate) >= edit.index && static_cast<std::uint64_t>(coordinate) < delete_end) {
        continue;  // The formula anchor itself was deleted.
      }
      if (static_cast<std::uint64_t>(coordinate) >= delete_end) {
        coordinate -= edit.count;
      }
    } else if (coordinate >= edit.index) {
      const std::uint64_t moved = static_cast<std::uint64_t>(coordinate) + edit.count;
      if (moved >= bound) {
        continue;  // The formula anchor moved outside the grid.
      }
      coordinate = static_cast<std::uint32_t>(moved);
    }

    BlockedSpillFootprint moved = footprint;
    moved.anchor_row = edit.row_axis ? coordinate : footprint.anchor_row;
    moved.anchor_col = edit.row_axis ? footprint.anchor_col : coordinate;
    shifted.emplace(CellAddress{moved.anchor_row, moved.anchor_col}, moved);
  }
  spill_table_->blocked_by_anchor = std::move(shifted);
}

void Sheet::clear_spill(std::uint32_t anchor_row, std::uint32_t anchor_col) noexcept {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  clear_spill_locked(anchor_row, anchor_col);
}

void Sheet::clear_all_spills() noexcept {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  clear_committed_spills_locked();
  if (spill_table_ != nullptr && !spill_table_->blocked_by_anchor.empty()) {
    spill_table_->blocked_by_anchor.clear();
    cell_enumeration_revision_.bump();
  }
}

void Sheet::clear_spill_locked(std::uint32_t anchor_row, std::uint32_t anchor_col) noexcept {
  if (spill_table_ == nullptr) {
    return;
  }
  const CellAddress anchor_addr{anchor_row, anchor_col};
  const auto region_it = spill_table_->by_anchor.find(anchor_addr);
  const auto blocked_it = spill_table_->blocked_by_anchor.find(anchor_addr);
  if (region_it == spill_table_->by_anchor.end() && blocked_it == spill_table_->blocked_by_anchor.end()) {
    return;
  }
  if (region_it != spill_table_->by_anchor.end()) {
    spill_table_->erase_region(region_it);
  }
  if (blocked_it != spill_table_->blocked_by_anchor.end()) {
    spill_table_->blocked_by_anchor.erase(blocked_it);
  }
  cell_enumeration_revision_.bump();
}

void Sheet::clear_committed_spills_locked() noexcept {
  if (spill_table_ == nullptr || spill_table_->by_anchor.empty()) {
    return;
  }
  spill_table_->clear_regions();
  cell_enumeration_revision_.bump();
}

}  // namespace formulon
