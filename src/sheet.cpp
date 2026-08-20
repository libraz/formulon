//
// Out-of-line implementation of the row-sparse, column-dense cell store
// owned by `Sheet` and the heap-owned spill-region table. See `sheet.h` for
// the storage-layer and spill API contracts.

#include "sheet.h"

#include <algorithm>
#include <cassert>
#include <cstddef>
#include <cstdint>
#include <functional>
#include <memory>
#include <mutex>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_types.h"
#include "io/a1_ref.h"
#include "pivot/pivot_table.h"
#include "utils/a1_column.h"
#include "utils/arena.h"
#include "utils/resource_budget.h"
#include "value.h"

namespace formulon {

// ---------------------------------------------------------------------------
// Private spill-table layout
// ---------------------------------------------------------------------------
//
// The table is keyed only by anchor cell and owns each region's payload.
// Phantom lookup scans these rectangles rather than retaining one hash-map
// entry per spilled cell: a 1,000 x 100 spill is one region, not 99,999
// duplicate reverse-index nodes. Iteration order is undefined; consumers
// that need a deterministic order must sort externally.
struct SpillTable {
  std::unordered_map<CellAddress, SpillRegion, CellAddressHash> by_anchor;
  // A failed dynamic-array commit leaves the anchor's attempted rectangle
  // here so later writes/removals inside that rectangle can re-dirty the
  // producer without requiring the user to touch the formula again.
  std::unordered_map<CellAddress, BlockedSpillFootprint, CellAddressHash> blocked_by_anchor;
};

namespace {

/// True when two half-open rectangles share at least one coordinate.
///
/// Every rectangle in this file arrives as an origin plus an extent, and the
/// ends are widened to 64 bits so a rectangle touching the last row or column
/// cannot wrap. Written once because the spill table, the merge list, the
/// blocked-footprint table and the bulk read all need the same test, and a
/// hand-rolled copy per caller is a place for one comparison to drift.
bool RectsIntersect(std::uint64_t a_row_begin, std::uint64_t a_row_end, std::uint64_t a_col_begin,
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

}  // namespace

// ---------------------------------------------------------------------------
// Special members (must be defined here where SpillTable is complete).
// ---------------------------------------------------------------------------

Sheet::Sheet(std::string name) : name_(std::move(name)), spill_mutex_(std::make_unique<std::mutex>()) {}

// Defaulting these out-of-line keeps the incomplete SpillTable deleter out of
// the header while ensuring new Sheet metadata participates in moves without
// a hand-maintained member list.
Sheet::Sheet(Sheet&&) noexcept = default;
Sheet& Sheet::operator=(Sheet&&) noexcept = default;

Sheet::~Sheet() = default;

void Sheet::add_pivot_table(std::unique_ptr<pivot::PivotTable> table) {
  if (table == nullptr) {
    return;
  }
  pivot_tables_.push_back(std::move(table));
}

const Cell& RowCells::blank() noexcept {
  // Shared read-only stand-in for a column the row never materialised.
  // `operator[]` hands it out for the leading gap, so index-based scans see a
  // default cell exactly where a dense vector would have held one.
  static const Cell kBlank;
  return kBlank;
}

Cell& RowCells::ensure(std::uint32_t col) {
  if (run_.empty()) {
    first_col_ = col;
    run_.resize(1U);
    return run_.front();
  }
  if (col < first_col_) {
    // Extending to the left re-seats the run. `Cell` is move-only, so build
    // the wider run and move the existing slots into place; the heap-stable
    // `cached_text_owned` payload each cell owns survives the move.
    const std::size_t shift = static_cast<std::size_t>(first_col_) - col;
    std::vector<Cell> grown;
    grown.resize(shift + run_.size());
    std::move(run_.begin(), run_.end(), grown.begin() + static_cast<std::ptrdiff_t>(shift));
    run_ = std::move(grown);
    first_col_ = col;
    return run_.front();
  }
  const std::size_t index = static_cast<std::size_t>(col) - first_col_;
  if (index >= run_.size()) {
    run_.resize(index + 1U);
  }
  return run_[index];
}

namespace {

// Returns `value` with any Text payload copied into `arena`.
//
// A Text `Value` read out of sheet storage is a view into bytes the sheet
// owns and frees: a cell's `cached_text_owned`, replaced by the next cached
// write, or a spill region's `owned_strings`, freed when the region is
// cleared. Re-homing the bytes under the same lock that read them is what
// lets the copy leave the critical section; the caller's arena decides how
// long it stays readable. Non-Text values carry no pointer.
Value AdoptText(const Value& value, Arena& arena) {
  if (!value.is_text()) {
    return value;
  }
  return Value::text(arena.intern(value.as_text()));
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

void Sheet::set_cell_value(std::uint32_t row, std::uint32_t col, Value v) {
  // Bounds checks are advisory: callers above this layer (parser, OOXML
  // reader) own coordinate validation. A debug assert catches programming
  // errors without imposing a release-mode branch.
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);

  // A literal write can replace either the anchor or a phantom. In both
  // cases the complete region must disappear before updating storage.
  const CellAddress address{row, col};
  if (spill_table_ != nullptr &&
      (spill_table_->by_anchor.find(address) != spill_table_->by_anchor.end() ||
       spill_table_->blocked_by_anchor.find(address) != spill_table_->blocked_by_anchor.end())) {
    clear_spill_locked(row, col);
  } else if (const SpillRegion* covering = spill_region_covering_locked(row, col); covering != nullptr) {
    clear_spill_locked(covering->anchor_row, covering->anchor_col);
  }

  RowCells& row_cells = rows_[row];
  Cell& slot = row_cells.ensure(col);
  slot.formula_text.clear();
  slot.phonetic_text.clear();
  slot.cached_value = v;
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_text(std::uint32_t row, std::uint32_t col, std::string_view text) {
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  const CellAddress address{row, col};
  if (spill_table_ != nullptr &&
      (spill_table_->by_anchor.find(address) != spill_table_->by_anchor.end() ||
       spill_table_->blocked_by_anchor.find(address) != spill_table_->blocked_by_anchor.end())) {
    clear_spill_locked(row, col);
  } else if (const SpillRegion* covering = spill_region_covering_locked(row, col); covering != nullptr) {
    clear_spill_locked(covering->anchor_row, covering->anchor_col);
  }

  RowCells& row_cells = rows_[row];
  Cell& slot = row_cells.ensure(col);
  slot.formula_text.clear();
  slot.phonetic_text.clear();
  auto owned = std::make_unique<std::string>(text);
  slot.cached_value = Value::text(*owned);
  slot.cached_text_owned = std::move(owned);
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_formula(std::uint32_t row, std::uint32_t col, std::string formula) {
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);

  // Formula replacement can likewise target an anchor or a phantom.
  const CellAddress address{row, col};
  if (spill_table_ != nullptr &&
      (spill_table_->by_anchor.find(address) != spill_table_->by_anchor.end() ||
       spill_table_->blocked_by_anchor.find(address) != spill_table_->blocked_by_anchor.end())) {
    clear_spill_locked(row, col);
  } else if (const SpillRegion* covering = spill_region_covering_locked(row, col); covering != nullptr) {
    clear_spill_locked(covering->anchor_row, covering->anchor_col);
  }

  RowCells& row_cells = rows_[row];
  Cell& slot = row_cells.ensure(col);
  slot.formula_text = std::move(formula);
  slot.phonetic_text.clear();
  slot.cached_value = Value::blank();
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_cached_value(std::uint32_t row, std::uint32_t col, Value v) {
  // Bounds checks are advisory: callers above this layer (parser, OOXML
  // reader, recalc engine) own coordinate validation. A debug assert
  // catches programming errors without imposing a release-mode branch.
  assert(row < kMaxRows && col < kMaxCols);

  // Cached-value updates do NOT trigger spill invalidation: the recalc
  // engine writes the post-evaluation result of a formula cell, and any
  // structural change (formula edit, literal write into a phantom) is
  // already routed through `set_cell_formula` / `set_cell_value` which
  // handle invalidation separately. Letting cached-value updates bypass
  // the spill table also keeps the spill anchor's `cached_value`
  // synchronised with `commit_spill` (which sets it to `cells[0]`).
  //
  // Locks `spill_mutex_` because it writes `rows_`, which the spill path
  // and concurrent observers also touch. The scheduler calls this under its
  // own `write_mutex`; the outer-to-inner lock ordering (`write_mutex` then
  // `spill_mutex_`) is preserved and never inverted.
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  RowCells& row_cells = rows_[row];
  Cell& slot = row_cells.ensure(col);

  if (v.is_text()) {
    // Deep-copy the Text payload into a heap-stable `std::string` owned by
    // the Cell so the stored `cached_value` no longer references whatever
    // buffer the caller used (typically the recalc engine's per-evaluation
    // `Arena`, which is about to be reset). Allocating a fresh
    // `unique_ptr<std::string>` per write keeps the bytes at a fixed heap
    // address that survives the Cell being relocated by `Sheet`'s row-
    // vector growth and row-map rehash paths — a bare `std::string` would
    // relocate its inline (SSO) bytes on every Cell move and dangle the
    // `string_view` we are about to store.
    //
    // Construct from `v.as_text()` directly: even when the caller passes
    // the cell's own current cached text back through this API, the new
    // `std::string` allocates its own buffer before we replace
    // `cached_text_owned`, so the source bytes remain valid for the
    // duration of the construction.
    auto owned = std::make_unique<std::string>(v.as_text());
    const std::string_view view(*owned);
    slot.cached_text_owned = std::move(owned);
    slot.cached_value = Value::text(view);
  } else {
    // Non-Text path: keep the trivial assignment. The previous
    // `cached_text_owned` (if any) is intentionally retained: the new
    // `cached_value` does not reference it, and freeing it eagerly would
    // produce no visible win. The next Text write replaces the pointer
    // unconditionally.
    slot.cached_value = v;
  }
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_cached_value_borrowed(std::uint32_t row, std::uint32_t col, Value v) {
  assert(row < kMaxRows && col < kMaxCols);

  // This is the reader-only counterpart to `set_cell_cached_value`: the
  // workbook-owned shared-string deque keeps Text payloads alive for the
  // workbook lifetime, so duplicating a shared value per cell would turn a
  // compact SST into O(number of cells) storage. Clear any evaluator-owned
  // backing allocation before installing the borrowed view.
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  RowCells& row_cells = rows_[row];
  Cell& slot = row_cells.ensure(col);
  slot.cached_text_owned.reset();
  slot.cached_value = v;
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_phonetic(std::uint32_t row, std::uint32_t col, std::string_view phonetic) {
  // Bounds checks are advisory: callers above this layer (OOXML reader)
  // own coordinate validation. A debug assert catches programming errors
  // without imposing a release-mode branch.
  assert(row < kMaxRows && col < kMaxCols);

  // Phonetic-annotation writes are not structural mutations of the cell
  // value — they parallel the surface text. Skip the spill-invalidation
  // dance that `set_cell_value` performs; the OOXML reader writes the
  // surface value first via `set_cell_cached_value`, and only the
  // post-hoc phonetic copy lands here.
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  RowCells& row_cells = rows_[row];
  Cell& slot = row_cells.ensure(col);
  slot.phonetic_text.assign(phonetic.begin(), phonetic.end());
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_xf_index(std::uint32_t row, std::uint32_t col, std::uint32_t xf_index) {
  assert(row < kMaxRows && col < kMaxCols);
  // Formatting is orthogonal to a dynamic-array spill. In particular, a
  // style write into a spill phantom must not clear the region as a literal
  // write would. It can still grow `rows_`, so share the sheet mutation lock
  // used by the spill and cached-value paths.
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  RowCells& row_cells = rows_[row];
  Cell& slot = row_cells.ensure(col);
  slot.xf_index = xf_index;
  cell_enumeration_revision_.bump();
}

const Cell* Sheet::cell_at(std::uint32_t row, std::uint32_t col) const noexcept {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  return cell_at_locked(row, col);
}

const Cell* Sheet::cell_at_locked(std::uint32_t row, std::uint32_t col) const noexcept {
  const auto it = rows_.find(row);
  if (it == rows_.end()) {
    return nullptr;
  }
  return it->second.find(col);
}

bool Sheet::has_cell(std::uint32_t row, std::uint32_t col) const noexcept {
  return cell_at(row, col) != nullptr;
}

std::size_t Sheet::cell_count() const noexcept {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  std::size_t total = 0;
  for (const auto& kv : rows_) {
    total += kv.second.stored_count();
  }
  // Add dynamic-array spill phantoms that have no underlying stored slot. A
  // phantom may coincide with an implicitly default-constructed slot (created
  // when the row's run grew to cover a later column); such a coordinate is
  // already counted above, so only phantoms absent from `rows_` add to the
  // total. This keeps the count aligned with the flat enumeration exposed
  // through the C ABI, which surfaces phantoms via `resolve_cell_value`.
  //
  // Derived from each region's area rather than walked cell by cell: a spill
  // is one rectangle, and a whole-column one is 1,048,576 coordinates whose
  // individual hash lookups would run under this lock while every other
  // reader of the sheet waits.
  if (spill_table_ != nullptr) {
    for (const auto& kv : spill_table_->by_anchor) {
      const SpillRegion& region = kv.second;
      const std::uint64_t row_end = static_cast<std::uint64_t>(region.anchor_row) + region.rows;
      const std::uint64_t col_end = static_cast<std::uint64_t>(region.anchor_col) + region.cols;
      const std::size_t area = static_cast<std::size_t>(region.rows) * static_cast<std::size_t>(region.cols);
      const std::size_t materialised =
          materialised_cells_in_rect_locked(region.anchor_row, row_end, region.anchor_col, col_end);
      // Materialised slots inside the rectangle are already in `total`, so
      // only the remainder is new. `commit_spill` materialises the anchor, so
      // subtracting the materialised count also removes the anchor the
      // phantom set excludes; a region whose anchor slot is somehow absent
      // needs that exclusion applied by hand.
      total += area - materialised;
      if (cell_at_locked(region.anchor_row, region.anchor_col) == nullptr) {
        --total;
      }
    }
  }
  return total;
}

std::size_t Sheet::materialised_cells_in_rect_locked(std::uint32_t row_begin, std::uint64_t row_end,
                                                     std::uint32_t col_begin, std::uint64_t col_end) const noexcept {
  // A row's materialised slots are one contiguous run, so its contribution is
  // the length of the run's overlap with the column span — no per-column
  // probing.
  const auto overlap = [&](const RowCells& cells) noexcept -> std::size_t {
    if (cells.empty()) {
      return 0U;
    }
    const std::uint64_t first = std::max<std::uint64_t>(col_begin, cells.first_col());
    const std::uint64_t last = std::min<std::uint64_t>(col_end, cells.size());
    return first < last ? static_cast<std::size_t>(last - first) : 0U;
  };

  // Same choice `footprint_holds_occupied_cell_locked` makes: whichever of
  // the stored rows and the rectangle's rows is the smaller set.
  std::size_t total = 0;
  if (static_cast<std::uint64_t>(rows_.size()) <= row_end - row_begin) {
    for (const auto& [row, cells] : rows_) {
      if (static_cast<std::uint64_t>(row) < row_begin || static_cast<std::uint64_t>(row) >= row_end) {
        continue;
      }
      total += overlap(cells);
    }
    return total;
  }
  for (std::uint64_t row = row_begin; row < row_end; ++row) {
    const auto it = rows_.find(static_cast<std::uint32_t>(row));
    if (it != rows_.end()) {
      total += overlap(it->second);
    }
  }
  return total;
}

std::optional<Sheet::PopulatedExtent> Sheet::populated_extent(std::uint32_t first_row, std::uint32_t first_col,
                                                              std::uint32_t last_row,
                                                              std::uint32_t last_col) const noexcept {
  if (!rect_in_grid(first_row, first_col, last_row, last_col)) {
    return std::nullopt;
  }

  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  PopulatedExtent extent;
  bool any = false;
  const auto include = [&](std::uint32_t row, std::uint32_t col) {
    if (!any) {
      extent.first_row = extent.last_row = row;
      extent.first_col = extent.last_col = col;
      any = true;
      return;
    }
    extent.first_row = std::min(extent.first_row, row);
    extent.first_col = std::min(extent.first_col, col);
    extent.last_row = std::max(extent.last_row, row);
    extent.last_col = std::max(extent.last_col, col);
  };

  // RowCells stores an absolute-origin run. Restrict the scan to the run's
  // materialised interval and inspect only non-blank/formula slots; leading
  // gaps are not represented and default-constructed slots are not content.
  for (const auto& [row, cells] : rows_) {
    if (row < first_row || row > last_row || cells.empty()) {
      continue;
    }
    const std::size_t begin = std::max<std::size_t>(first_col, cells.first_col());
    const std::size_t end = std::min<std::size_t>(static_cast<std::size_t>(last_col) + 1U, cells.size());
    if (begin >= end) {
      continue;
    }
    for (std::size_t col = begin; col < end; ++col) {
      const Cell& cell = cells[col];
      if (!cell.formula_text.empty() || !cell.cached_value.is_blank()) {
        include(row, static_cast<std::uint32_t>(col));
      }
    }
  }

  // A committed spill is represented by one rectangle, so fold its clipped
  // intersection into the result without walking its phantom payload.
  if (spill_table_ != nullptr) {
    for (const auto& [unused, region] : spill_table_->by_anchor) {
      (void)unused;
      const std::uint64_t region_last_row = static_cast<std::uint64_t>(region.anchor_row) + region.rows - 1U;
      const std::uint64_t region_last_col = static_cast<std::uint64_t>(region.anchor_col) + region.cols - 1U;
      if (region.anchor_row > last_row || region.anchor_col > last_col || region_last_row < first_row ||
          region_last_col < first_col) {
        continue;
      }
      const std::uint32_t clipped_first_row = std::max(first_row, region.anchor_row);
      const std::uint32_t clipped_first_col = std::max(first_col, region.anchor_col);
      const std::uint32_t clipped_last_row = std::min(last_row, static_cast<std::uint32_t>(region_last_row));
      const std::uint32_t clipped_last_col = std::min(last_col, static_cast<std::uint32_t>(region_last_col));
      include(clipped_first_row, clipped_first_col);
      include(clipped_last_row, clipped_last_col);
    }
  }

  if (!any) {
    return std::nullopt;
  }
  return extent;
}

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
  for (const auto& [unused, region] : spill_table_->by_anchor) {
    (void)unused;
    if (!RectIntersectsSpan(region, row, static_cast<std::uint64_t>(row) + 1U, col,
                            static_cast<std::uint64_t>(col) + 1U)) {
      continue;
    }
    return BlockedSpillFootprint{region.anchor_row, region.anchor_col, region.rows, region.cols};
  }
  return std::nullopt;
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

void Sheet::add_merge(MergeRange merge) {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  merges_.push_back(merge);
}

std::vector<MergeRange> Sheet::remove_merges_intersecting(MergeRange range) {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  std::vector<MergeRange> removed;
  for (auto it = merges_.begin(); it != merges_.end();) {
    if (it->last_row < range.first_row || range.last_row < it->first_row || it->last_col < range.first_col ||
        range.last_col < it->first_col) {
      ++it;
      continue;
    }
    removed.push_back(*it);
    it = merges_.erase(it);
  }
  return removed;
}

bool Sheet::remove_merge_at(std::size_t index, MergeRange* removed) {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (index >= merges_.size()) {
    return false;
  }
  if (removed != nullptr) {
    *removed = merges_[index];
  }
  merges_.erase(merges_.begin() + static_cast<std::ptrdiff_t>(index));
  return true;
}

std::size_t Sheet::clear_merges() {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  const std::size_t removed = merges_.size();
  merges_.clear();
  return removed;
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
  for (const auto& entry : spill_table_->by_anchor) {
    const SpillRegion& region = entry.second;
    if (!RectIntersectsSpan(region, row, static_cast<std::uint64_t>(row) + 1U, col,
                            static_cast<std::uint64_t>(col) + 1U)) {
      continue;
    }
    if (row == region.anchor_row && col == region.anchor_col) {
      return nullptr;
    }
    return &region;
  }
  return nullptr;
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
  // `RecalcEngine`'s ordering-edge walk.
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
  // their own tables, so they already cost their table size rather than the
  // footprint's area; only the stored-cell sweep needed the sparse walk.
  //
  // A pre-existing region at this anchor is the producer's own spill. It is
  // ignored wholesale: ad-hoc evaluation is read-only and cannot clear it,
  // while commit_spill clears it before reaching this predicate.
  if (spill_table_ != nullptr) {
    for (const auto& entry : spill_table_->by_anchor) {
      const SpillRegion& region = entry.second;
      if (region.anchor_row == anchor_row && region.anchor_col == anchor_col) {
        continue;
      }
      if (RectIntersectsSpan(region, anchor_row, row_end, anchor_col, col_end)) {
        return SpillAdmission::kBlocked;
      }
    }
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

Value Sheet::resolve_cell_value(std::uint32_t row, std::uint32_t col) const noexcept {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (const SpillRegion* covering = spill_region_covering_locked(row, col); covering != nullptr) {
    const std::uint32_t r_off = row - covering->anchor_row;
    const std::uint32_t c_off = col - covering->anchor_col;
    const std::size_t index =
        static_cast<std::size_t>(r_off) * static_cast<std::size_t>(covering->cols) + static_cast<std::size_t>(c_off);
    return covering->cells[index];
  }
  if (const Cell* c = cell_at_locked(row, col); c != nullptr) {
    return c->cached_value;
  }
  return Value::blank();
}

void Sheet::read_formula_cell(std::uint32_t row, std::uint32_t col, CellRead& out) const {
  out.exists_ = false;
  out.is_text_ = false;
  out.formula_text_.clear();
  out.text_payload_.clear();
  out.value_ = Value::blank();
  if (!coord_in_grid(row, col)) {
    return;
  }

  // Everything below runs inside the one critical section, and nothing but
  // copies leaves it.
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  const Cell* cell = cell_at_locked(row, col);
  out.exists_ = cell != nullptr;

  Value value = Value::blank();
  if (cell != nullptr && !cell->formula_text.empty()) {
    // Formula cell. A formula cell is never a spill phantom — it is either
    // an ordinary cell or a spill anchor, and `commit_spill` keeps an
    // anchor's `cached_value` equal to its region's first cell.
    out.formula_text_ = cell->formula_text;
    value = cell->cached_value;
  } else if (const SpillRegion* covering = spill_region_covering_locked(row, col); covering != nullptr) {
    const std::size_t index = static_cast<std::size_t>(row - covering->anchor_row) * covering->cols +
                              static_cast<std::size_t>(col - covering->anchor_col);
    value = covering->cells[index];
  } else if (cell != nullptr) {
    value = cell->cached_value;
  }

  // A Text payload is a view into storage the writer owns; copy the bytes
  // so the reading thread stops depending on that allocation's lifetime.
  if (value.is_text()) {
    out.is_text_ = true;
    out.text_payload_ = value.as_text();
  } else {
    out.value_ = value;
  }
}

void Sheet::read_range(std::uint32_t first_row, std::uint32_t last_row, std::uint32_t first_col, std::uint32_t last_col,
                       Arena& text_arena, std::vector<Value>& out, std::vector<std::size_t>& formula_indices) const {
  if (first_row > last_row || first_col > last_col || last_row >= kMaxRows || last_col >= kMaxCols) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);

  // Narrow the spill table to the regions this rectangle can actually reach,
  // once. `spill_region_covering_locked` is a linear scan of every registered
  // anchor, and calling it per coordinate makes the read cost area x table
  // size — the opposite of the amortisation this method exists for.
  std::vector<const SpillRegion*> regions;
  if (spill_table_ != nullptr) {
    const std::uint64_t row_end = static_cast<std::uint64_t>(last_row) + 1U;
    const std::uint64_t col_end = static_cast<std::uint64_t>(last_col) + 1U;
    for (const auto& entry : spill_table_->by_anchor) {
      if (RectIntersectsSpan(entry.second, first_row, row_end, first_col, col_end)) {
        regions.push_back(&entry.second);
      }
    }
  }
  const auto covering = [&regions](std::uint32_t row, std::uint32_t col) noexcept -> const SpillRegion* {
    for (const SpillRegion* region : regions) {
      if (!RectIntersectsSpan(*region, row, static_cast<std::uint64_t>(row) + 1U, col,
                              static_cast<std::uint64_t>(col) + 1U)) {
        continue;
      }
      // The anchor reads as its own stored cell, matching the scalar path.
      return (row == region->anchor_row && col == region->anchor_col) ? nullptr : region;
    }
    return nullptr;
  };

  for (std::uint32_t row = first_row; row <= last_row; ++row) {
    const auto row_it = rows_.find(row);
    const RowCells* stored = row_it != rows_.end() ? &row_it->second : nullptr;
    for (std::uint32_t col = first_col; col <= last_col; ++col) {
      const Cell* cell = stored != nullptr ? stored->find(col) : nullptr;
      // Order matters and mirrors the scalar path: a stored formula wins over
      // a covering spill region, which in turn wins over a stored literal.
      if (cell != nullptr && !cell->formula_text.empty()) {
        formula_indices.push_back(out.size());
        out.push_back(AdoptText(cell->cached_value, text_arena));
        continue;
      }
      if (const SpillRegion* region = covering(row, col); region != nullptr) {
        const std::size_t index = static_cast<std::size_t>(row - region->anchor_row) * region->cols +
                                  static_cast<std::size_t>(col - region->anchor_col);
        out.push_back(AdoptText(region->cells[index], text_arena));
        continue;
      }
      out.push_back(cell != nullptr ? AdoptText(cell->cached_value, text_arena) : Value::blank());
    }
  }
}

bool Sheet::read_spill_region_at_anchor(std::uint32_t row, std::uint32_t col, Arena& text_arena,
                                        std::vector<Value>& out_cells, std::uint32_t* out_rows,
                                        std::uint32_t* out_cols) const {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (spill_table_ == nullptr) {
    return false;
  }
  const auto it = spill_table_->by_anchor.find(CellAddress{row, col});
  if (it == spill_table_->by_anchor.end()) {
    return false;
  }
  const SpillRegion& region = it->second;
  out_cells.reserve(out_cells.size() + region.cells.size());
  for (const Value& value : region.cells) {
    out_cells.push_back(AdoptText(value, text_arena));
  }
  if (out_rows != nullptr) {
    *out_rows = region.rows;
  }
  if (out_cols != nullptr) {
    *out_cols = region.cols;
  }
  return true;
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
  const auto inserted = spill_table_->by_anchor.emplace(anchor_addr, std::move(region));
  assert(inserted.second && "commit_spill: anchor entry already present after clear");
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

namespace {

// Shifts every entry in `items` whose anchor lies on or past `index` by the
// insert / delete rule encoded in `count` and `is_delete`. Entries whose
// anchor falls inside the deleted interval, or whose insert would push it
// past `bound`, are removed. `field` names the axis coordinate, so one
// instantiation per anchor type serves both axes.
template <typename Anchor>
void ShiftAnchored(std::vector<Anchor>& items, std::uint32_t Anchor::*field, std::uint32_t index, std::uint32_t count,
                   bool is_delete, std::uint32_t bound) {
  std::vector<Anchor> retained;
  retained.reserve(items.size());
  for (Anchor& item : items) {
    if (item.*field < index) {
      retained.push_back(std::move(item));
      continue;
    }
    if (is_delete) {
      if (item.*field < index + count) {
        continue;  // Anchor inside deleted interval; drop the entry.
      }
      item.*field -= count;
    } else {
      // Insert. Anchors at or past `index` shift forward; entries pushed
      // past the sheet bound are dropped.
      const std::uint64_t shifted = static_cast<std::uint64_t>(item.*field) + count;
      if (shifted >= bound) {
        continue;
      }
      item.*field = static_cast<std::uint32_t>(shifted);
    }
    retained.push_back(std::move(item));
  }
  items = std::move(retained);
}

template <typename Anchor>
void ShiftRowAnchored(std::vector<Anchor>& items, std::uint32_t row, std::uint32_t count, bool is_delete) {
  ShiftAnchored(items, &Anchor::row, row, count, is_delete, Sheet::kMaxRows);
}

template <typename Anchor>
void ShiftColAnchored(std::vector<Anchor>& items, std::uint32_t col, std::uint32_t count, bool is_delete) {
  ShiftAnchored(items, &Anchor::col, col, count, is_delete, Sheet::kMaxCols);
}

// Rectangular merge / validation range shifter along one axis.
// Spans that fall entirely inside the deleted interval are dropped;
// spans that straddle the deletion are clamped so the survivors stay
// contiguous (Excel's "shrink the merge" behaviour). Inserts that would
// push `last` past `bound - 1` clamp to the sheet bound.
void ShiftSpan(std::uint32_t& first, std::uint32_t& last, std::uint32_t index, std::uint32_t count, bool is_delete,
               std::uint32_t bound, bool* out_drop) {
  *out_drop = false;
  if (is_delete) {
    const std::uint32_t del_end = index + count;  // exclusive
    // Both endpoints below the deletion: unchanged.
    if (last < index) {
      return;
    }
    // Both endpoints inside the deletion: drop the range entirely.
    if (first >= index && last < del_end) {
      *out_drop = true;
      return;
    }
    // Split shifts depending on which endpoints fall inside.
    if (first < index && last >= index && last < del_end) {
      // Trailing endpoint inside the deletion; clamp to index-1.
      last = index - 1U;
      return;
    }
    if (first >= index && first < del_end && last >= del_end) {
      // Leading endpoint inside the deletion; clamp to the line after the
      // deletion (which after the shift becomes `index`).
      first = index;
      last -= count;
      return;
    }
    if (first < index && last >= del_end) {
      // Range straddles the entire deletion: shrink by `count`. The
      // leading endpoint stays put; the trailing one shifts back.
      last -= count;
      return;
    }
    // Both endpoints past the deletion: shift back.
    first -= count;
    last -= count;
    return;
  }
  // Insert. Endpoints at or past `index` shift forward; clamp to bound.
  if (last < index) {
    return;  // Both endpoints below the insert; unchanged.
  }
  auto shift_one = [count, bound](std::uint32_t value) -> std::uint32_t {
    const std::uint64_t shifted = static_cast<std::uint64_t>(value) + count;
    if (shifted >= bound) {
      return bound - 1U;
    }
    return static_cast<std::uint32_t>(shifted);
  };
  if (first >= index) {
    first = shift_one(first);
  }
  last = shift_one(last);
}

void ShiftRowRange(MergeRange& range, std::uint32_t row, std::uint32_t count, bool is_delete, bool* out_drop) {
  ShiftSpan(range.first_row, range.last_row, row, count, is_delete, Sheet::kMaxRows, out_drop);
}

void ShiftColRange(MergeRange& range, std::uint32_t col, std::uint32_t count, bool is_delete, bool* out_drop) {
  ShiftSpan(range.first_col, range.last_col, col, count, is_delete, Sheet::kMaxCols, out_drop);
}

void ShiftRangeList(std::vector<MergeRange>& ranges, std::uint32_t index, std::uint32_t count, bool is_delete,
                    bool row_axis) {
  std::vector<MergeRange> retained;
  retained.reserve(ranges.size());
  for (MergeRange& range : ranges) {
    bool drop = false;
    if (row_axis) {
      ShiftRowRange(range, index, count, is_delete, &drop);
    } else {
      ShiftColRange(range, index, count, is_delete, &drop);
    }
    if (drop) {
      continue;
    }
    retained.push_back(range);
  }
  ranges = std::move(retained);
}

// Hyperlinks use the same inclusive rectangle semantics as merge ranges.
// Keeping this conversion next to ShiftRangeList is intentional: row/column
// insertions and deletions must apply the exact same shift, shrink, clamp and
// drop rules to both metadata shapes.
void ShiftHyperlinkList(std::vector<Hyperlink>& hyperlinks, std::uint32_t index, std::uint32_t count, bool is_delete,
                        bool row_axis) {
  std::vector<Hyperlink> retained;
  retained.reserve(hyperlinks.size());
  for (Hyperlink& hyperlink : hyperlinks) {
    MergeRange range{hyperlink.row, hyperlink.col, hyperlink.last_row, hyperlink.last_col};
    bool drop = false;
    if (row_axis) {
      ShiftRowRange(range, index, count, is_delete, &drop);
    } else {
      ShiftColRange(range, index, count, is_delete, &drop);
    }
    if (drop) {
      continue;
    }
    hyperlink.row = range.first_row;
    hyperlink.col = range.first_col;
    hyperlink.last_row = range.last_row;
    hyperlink.last_col = range.last_col;
    retained.push_back(std::move(hyperlink));
  }
  hyperlinks = std::move(retained);
}

void ShiftConditionalFormats(std::vector<cf::ConditionalFormat>& formats, std::uint32_t index, std::uint32_t count,
                             bool is_delete, bool row_axis) {
  std::vector<cf::ConditionalFormat> retained_formats;
  retained_formats.reserve(formats.size());
  for (cf::ConditionalFormat& format : formats) {
    std::vector<cf::CFCellRange> retained;
    retained.reserve(format.sqref.size());
    for (cf::CFCellRange& range : format.sqref) {
      MergeRange shifted{range.first.row, range.first.col, range.last.row, range.last.col};
      bool drop = false;
      if (row_axis) {
        ShiftRowRange(shifted, index, count, is_delete, &drop);
      } else {
        ShiftColRange(shifted, index, count, is_delete, &drop);
      }
      if (!drop) {
        range.first = CellAddress{shifted.first_row, shifted.first_col};
        range.last = CellAddress{shifted.last_row, shifted.last_col};
        retained.push_back(std::move(range));
      }
    }
    format.sqref = std::move(retained);
    // A conditionalFormatting element without an sqref applies nowhere and
    // is invalid OOXML. Drop it when a row/column deletion consumed every
    // one of its ranges.
    if (!format.sqref.empty()) {
      retained_formats.push_back(std::move(format));
    }
  }
  formats = std::move(retained_formats);
}

void ShiftRowLayouts(std::vector<RowLayout>& rows, std::uint32_t index, std::uint32_t count, bool is_delete) {
  ShiftRowAnchored(rows, index, count, is_delete);
}

/// Renders `[first_row..last_row] x [first_col..last_col]` as an OOXML
/// `ref` rectangle, collapsing a single-cell rectangle to one address the
/// way Excel writes it. Returns false when a coordinate is outside the grid.
bool AppendRefRectangle(std::string& out, const MergeRange& rect) {
  const auto append_cell = [&out](std::uint32_t row, std::uint32_t col) {
    if (!a1::append_column_letters(out, col)) {
      return false;
    }
    out += std::to_string(static_cast<std::uint64_t>(row) + 1U);
    return true;
  };
  if (!append_cell(rect.first_row, rect.first_col)) {
    return false;
  }
  if (rect.first_row == rect.last_row && rect.first_col == rect.last_col) {
    return true;
  }
  out += ':';
  return append_cell(rect.last_row, rect.last_col);
}

/// Rewrites the `ref` rectangle of a raw `<autoFilter>` element through the
/// same span rules the modelled rectangles follow, and clears the element
/// when the edit consumed its whole rectangle.
///
/// The element is retained verbatim, so this is the one coordinate inside it
/// that moves. It is also the one that decides which cells the filter is
/// attached to: leaving it behind points the filter at whatever occupies the
/// old rectangle after the edit. The `colId` offsets on any `<filterColumn>`
/// children are relative to this rectangle's first column and are not
/// remapped — see the field's declaration for what that costs.
void ShiftAutoFilterRef(std::string& xml, std::uint32_t index, std::uint32_t count, bool is_delete, bool row_axis) {
  if (xml.empty()) {
    return;
  }
  // Confine the search to the start tag: only the `<autoFilter>` element
  // itself carries the rectangle.
  const std::size_t tag_end = xml.find('>');
  const std::size_t attr = xml.find("ref=\"");
  if (tag_end == std::string::npos || attr == std::string::npos || attr > tag_end) {
    return;
  }
  const std::size_t value_begin = attr + 5U;
  const std::size_t value_end = xml.find('"', value_begin);
  if (value_end == std::string::npos || value_end > tag_end) {
    return;
  }

  const std::string_view value(xml.data() + value_begin, value_end - value_begin);
  const std::size_t colon = value.find(':');
  MergeRange rect;
  if (!io::parse_a1_ref(value.substr(0, colon), &rect.first_row, &rect.first_col)) {
    return;  // Not a plain A1 rectangle; leave the element untouched.
  }
  if (colon == std::string_view::npos) {
    rect.last_row = rect.first_row;
    rect.last_col = rect.first_col;
  } else if (!io::parse_a1_ref(value.substr(colon + 1U), &rect.last_row, &rect.last_col)) {
    return;
  }
  if (rect.first_row > rect.last_row || rect.first_col > rect.last_col) {
    return;
  }

  bool drop = false;
  if (row_axis) {
    ShiftRowRange(rect, index, count, is_delete, &drop);
  } else {
    ShiftColRange(rect, index, count, is_delete, &drop);
  }
  if (drop) {
    // Every filtered cell was deleted, which is what Excel resolves by
    // removing the filter rather than by keeping an empty one.
    xml.clear();
    return;
  }
  std::string replacement;
  if (!AppendRefRectangle(replacement, rect)) {
    return;
  }
  xml.replace(value_begin, value_end - value_begin, replacement);
}

void ShiftColumnLayouts(std::vector<ColumnLayout>& columns, std::uint32_t index, std::uint32_t count, bool is_delete) {
  std::vector<ColumnLayout> retained;
  retained.reserve(columns.size());
  for (ColumnLayout& column : columns) {
    MergeRange shifted{0, column.first, 0, column.last};
    bool drop = false;
    ShiftColRange(shifted, index, count, is_delete, &drop);
    if (!drop) {
      column.first = shifted.first_col;
      column.last = shifted.last_col;
      retained.push_back(std::move(column));
    }
  }
  columns = std::move(retained);
}

void ShiftBreaks(std::vector<ManualBreak>& breaks, std::uint32_t index, std::uint32_t count, bool is_delete,
                 std::uint32_t bound) {
  std::vector<ManualBreak> retained;
  retained.reserve(breaks.size());
  for (ManualBreak& page_break : breaks) {
    if (page_break.id < index) {
      retained.push_back(std::move(page_break));
      continue;
    }
    if (is_delete) {
      if (page_break.id < index + count) {
        continue;
      }
      page_break.id -= count;
    } else {
      const std::uint64_t shifted = static_cast<std::uint64_t>(page_break.id) + count;
      if (shifted >= bound) {
        continue;
      }
      page_break.id = static_cast<std::uint32_t>(shifted);
    }
    retained.push_back(std::move(page_break));
  }
  breaks = std::move(retained);
}

void ShiftPivotAnchors(std::vector<std::unique_ptr<pivot::PivotTable>>& pivots, std::uint32_t index,
                       std::uint32_t count, bool is_delete, bool row_axis) {
  for (std::unique_ptr<pivot::PivotTable>& pivot : pivots) {
    if (pivot == nullptr) {
      continue;
    }
    std::uint32_t anchor = row_axis ? pivot->anchor_row() : pivot->anchor_col();
    if (anchor >= index) {
      if (is_delete) {
        anchor = anchor < index + count ? index : anchor - count;
      } else {
        const std::uint64_t shifted = static_cast<std::uint64_t>(anchor) + count;
        anchor = static_cast<std::uint32_t>(
            std::min<std::uint64_t>(shifted, row_axis ? Sheet::kMaxRows - 1U : Sheet::kMaxCols - 1U));
      }
    }
    pivot->set_anchor(row_axis ? anchor : pivot->anchor_row(), row_axis ? pivot->anchor_col() : anchor,
                      pivot->span_rows(), pivot->span_cols());
  }
}

}  // namespace

// Every sheet-attached structure derives its new coordinates from the one
// `StructuralEdit` description. This list is the enumeration: a structure
// added to `Sheet` and not added here does not follow a row/column edit.
void Sheet::shift_sheet_metadata(const StructuralEdit& edit) {
  const std::uint32_t index = edit.index;
  const std::uint32_t count = edit.count;
  const bool is_delete = edit.is_delete;
  const bool row_axis = edit.row_axis;
  if (row_axis) {
    ShiftHyperlinkList(hyperlinks_, index, count, is_delete, /*row_axis=*/true);
    ShiftRowAnchored(comments_, index, count, is_delete);
  } else {
    ShiftHyperlinkList(hyperlinks_, index, count, is_delete, /*row_axis=*/false);
    ShiftColAnchored(comments_, index, count, is_delete);
  }
  ShiftRangeList(merges_, index, count, is_delete, row_axis);
  for (DataValidation& dv : validations_) {
    ShiftRangeList(dv.ranges, index, count, is_delete, row_axis);
  }
  ShiftConditionalFormats(conditional_formats_, index, count, is_delete, row_axis);
  if (row_axis) {
    ShiftRowLayouts(layout_.row_overrides, index, count, is_delete);
    ShiftBreaks(print_settings_.manual_row_breaks, index, count, is_delete, Sheet::kMaxRows);
  } else {
    ShiftColumnLayouts(layout_.columns, index, count, is_delete);
    ShiftBreaks(print_settings_.manual_col_breaks, index, count, is_delete, Sheet::kMaxCols);
  }
  ShiftPivotAnchors(pivot_tables_, index, count, is_delete, row_axis);
  ShiftAutoFilterRef(auto_filter_xml_, index, count, is_delete, row_axis);
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

void Sheet::insert_rows(std::uint32_t row, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  // Walk populated rows in descending key order so a moved row never
  // collides with an existing row that still needs to move.
  std::vector<std::uint32_t> keys;
  keys.reserve(rows_.size());
  for (const auto& kv : rows_) {
    keys.push_back(kv.first);
  }
  std::sort(keys.begin(), keys.end(), std::greater<std::uint32_t>());
  for (std::uint32_t key : keys) {
    if (key < row) {
      continue;
    }
    const std::uint64_t shifted = static_cast<std::uint64_t>(key) + count;
    auto node = rows_.extract(key);
    if (shifted >= kMaxRows) {
      continue;  // Row pushed past sheet bound; drop the cells.
    }
    node.key() = static_cast<std::uint32_t>(shifted);
    rows_.insert(std::move(node));
  }
  clear_committed_spills_locked();
  const StructuralEdit edit{row, count, /*is_delete=*/false, /*row_axis=*/true};
  shift_blocked_spills_locked(edit);
  shift_sheet_metadata(edit);
  cell_enumeration_revision_.bump();
}

void Sheet::delete_rows(std::uint32_t row, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  // Drop cells inside the deletion interval, then shift trailing rows
  // up. Walk in ascending key order — every shifted destination key is
  // strictly less than its source so no collision can occur.
  std::vector<std::uint32_t> keys;
  keys.reserve(rows_.size());
  for (const auto& kv : rows_) {
    keys.push_back(kv.first);
  }
  std::sort(keys.begin(), keys.end());
  for (std::uint32_t key : keys) {
    if (key < row) {
      continue;
    }
    auto node = rows_.extract(key);
    if (key < row + count) {
      continue;  // Deleted row; drop the node.
    }
    node.key() = key - count;
    rows_.insert(std::move(node));
  }
  clear_committed_spills_locked();
  const StructuralEdit edit{row, count, /*is_delete=*/true, /*row_axis=*/true};
  shift_blocked_spills_locked(edit);
  shift_sheet_metadata(edit);
  cell_enumeration_revision_.bump();
}

void Sheet::insert_cols(std::uint32_t col, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  for (auto& kv : rows_) {
    RowCells& row = kv.second;
    if (row.empty() || col >= row.size()) {
      continue;
    }
    // Insertion entirely left of the run: the run keeps its contents and only
    // its origin moves. A run pushed wholly past the last column is dropped.
    if (col <= row.first_col()) {
      const std::uint64_t shifted = static_cast<std::uint64_t>(row.first_col()) + count;
      if (shifted >= kMaxCols) {
        row.mutable_run().clear();
        row.set_first_col(0U);
        continue;
      }
      row.set_first_col(static_cast<std::uint32_t>(shifted));
      // The tail may now overhang the sheet; drop what no longer fits.
      const std::size_t capacity = static_cast<std::size_t>(kMaxCols) - row.first_col();
      if (row.mutable_run().size() > capacity) {
        row.mutable_run().resize(capacity);
      }
      continue;
    }
    // Insertion point sits inside this row's run. Pad with default Cells at
    // the local index. Cells past `kMaxCols` are dropped wholesale.
    std::vector<Cell>& cells = row.mutable_run();
    const std::size_t local_col = static_cast<std::size_t>(col) - row.first_col();
    const std::size_t old_size = cells.size();
    const std::uint64_t new_size_unbounded = static_cast<std::uint64_t>(old_size) + count;
    const std::size_t new_size =
        static_cast<std::size_t>(std::min<std::uint64_t>(new_size_unbounded, kMaxCols - row.first_col()));
    cells.resize(new_size);
    // Move the tail to its new position. Walk from the back so source
    // and destination never overlap during the swap.
    const std::size_t tail_count = old_size - local_col;
    for (std::size_t i = 0; i < tail_count; ++i) {
      const std::size_t src = old_size - 1U - i;
      const std::size_t dst = src + count;
      if (dst >= new_size) {
        continue;  // Cell pushed past sheet bound; drop it.
      }
      cells[dst] = std::move(cells[src]);
    }
    // Clear the inserted slots so they read as default-blank.
    const std::size_t clear_end = std::min<std::size_t>(local_col + count, new_size);
    for (std::size_t i = local_col; i < clear_end; ++i) {
      cells[i] = Cell{};
    }
  }
  clear_committed_spills_locked();
  const StructuralEdit edit{col, count, /*is_delete=*/false, /*row_axis=*/false};
  shift_blocked_spills_locked(edit);
  shift_sheet_metadata(edit);
  cell_enumeration_revision_.bump();
}

void Sheet::delete_cols(std::uint32_t col, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  for (auto& kv : rows_) {
    RowCells& row = kv.second;
    if (row.empty() || col >= row.size()) {
      continue;
    }
    std::vector<Cell>& cells = row.mutable_run();
    const std::uint64_t band_end = static_cast<std::uint64_t>(col) + count;
    if (band_end <= row.first_col()) {
      // The deleted band lies entirely left of the run: only the origin moves.
      row.set_first_col(row.first_col() - count);
      continue;
    }
    // The band reaches into the run. Everything from the run's start up to the
    // band's end goes; what remains starts at `col`.
    const std::size_t local_first = col > row.first_col() ? static_cast<std::size_t>(col) - row.first_col() : 0U;
    const std::size_t local_last =
        static_cast<std::size_t>(std::min<std::uint64_t>(band_end - row.first_col(), cells.size()));
    cells.erase(cells.begin() + static_cast<std::ptrdiff_t>(local_first),
                cells.begin() + static_cast<std::ptrdiff_t>(local_last));
    if (col < row.first_col()) {
      row.set_first_col(col);
    }
  }
  clear_committed_spills_locked();
  const StructuralEdit edit{col, count, /*is_delete=*/true, /*row_axis=*/false};
  shift_blocked_spills_locked(edit);
  shift_sheet_metadata(edit);
  cell_enumeration_revision_.bump();
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
    spill_table_->by_anchor.erase(region_it);
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
  spill_table_->by_anchor.clear();
  cell_enumeration_revision_.bump();
}

}  // namespace formulon
