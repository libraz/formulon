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
#include "pivot/pivot_table.h"
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
};

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

namespace {

// Grows `row_cells` so that index `col` is addressable, padding with
// default-constructed cells, and returns a reference to the slot at `col`.
Cell& EnsureSlot(std::vector<Cell>& row_cells, std::uint32_t col) {
  const std::size_t needed = static_cast<std::size_t>(col) + 1U;
  if (row_cells.size() < needed) {
    row_cells.resize(needed);
  }
  return row_cells[col];
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
  if (spill_table_ != nullptr && spill_table_->by_anchor.find(address) != spill_table_->by_anchor.end()) {
    clear_spill_locked(row, col);
  } else if (const SpillRegion* covering = spill_region_covering_locked(row, col); covering != nullptr) {
    clear_spill_locked(covering->anchor_row, covering->anchor_col);
  }

  std::vector<Cell>& row_cells = rows_[row];
  Cell& slot = EnsureSlot(row_cells, col);
  slot.formula_text.clear();
  slot.phonetic_text.clear();
  slot.cached_value = v;
}

void Sheet::set_cell_text(std::uint32_t row, std::uint32_t col, std::string_view text) {
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  const CellAddress address{row, col};
  if (spill_table_ != nullptr && spill_table_->by_anchor.find(address) != spill_table_->by_anchor.end()) {
    clear_spill_locked(row, col);
  } else if (const SpillRegion* covering = spill_region_covering_locked(row, col); covering != nullptr) {
    clear_spill_locked(covering->anchor_row, covering->anchor_col);
  }

  std::vector<Cell>& row_cells = rows_[row];
  Cell& slot = EnsureSlot(row_cells, col);
  slot.formula_text.clear();
  slot.phonetic_text.clear();
  auto owned = std::make_unique<std::string>(text);
  slot.cached_value = Value::text(*owned);
  slot.cached_text_owned = std::move(owned);
}

void Sheet::set_cell_formula(std::uint32_t row, std::uint32_t col, std::string formula) {
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);

  // Formula replacement can likewise target an anchor or a phantom.
  const CellAddress address{row, col};
  if (spill_table_ != nullptr && spill_table_->by_anchor.find(address) != spill_table_->by_anchor.end()) {
    clear_spill_locked(row, col);
  } else if (const SpillRegion* covering = spill_region_covering_locked(row, col); covering != nullptr) {
    clear_spill_locked(covering->anchor_row, covering->anchor_col);
  }

  std::vector<Cell>& row_cells = rows_[row];
  Cell& slot = EnsureSlot(row_cells, col);
  slot.formula_text = std::move(formula);
  slot.phonetic_text.clear();
  slot.cached_value = Value::blank();
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
  std::vector<Cell>& row_cells = rows_[row];
  Cell& slot = EnsureSlot(row_cells, col);

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
  std::vector<Cell>& row_cells = rows_[row];
  Cell& slot = EnsureSlot(row_cells, col);
  slot.phonetic_text.assign(phonetic.begin(), phonetic.end());
}

void Sheet::set_cell_xf_index(std::uint32_t row, std::uint32_t col, std::uint32_t xf_index) {
  assert(row < kMaxRows && col < kMaxCols);
  // Formatting is orthogonal to a dynamic-array spill. In particular, a
  // style write into a spill phantom must not clear the region as a literal
  // write would. It can still grow `rows_`, so share the sheet mutation lock
  // used by the spill and cached-value paths.
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  std::vector<Cell>& row_cells = rows_[row];
  Cell& slot = EnsureSlot(row_cells, col);
  slot.xf_index = xf_index;
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
  const std::vector<Cell>& row_cells = it->second;
  if (col >= row_cells.size()) {
    return nullptr;
  }
  return &row_cells[col];
}

bool Sheet::has_cell(std::uint32_t row, std::uint32_t col) const noexcept {
  return cell_at(row, col) != nullptr;
}

std::size_t Sheet::cell_count() const noexcept {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  std::size_t total = 0;
  for (const auto& kv : rows_) {
    total += kv.second.size();
  }
  // Add dynamic-array spill phantoms that have no underlying stored slot. A
  // phantom may coincide with an implicitly default-constructed slot (created
  // when the row vector grew to cover a later column); such a coordinate is
  // already counted above, so only phantoms absent from `rows_` add to the
  // total. This keeps the count aligned with the flat enumeration exposed
  // through the C ABI, which surfaces phantoms via `resolve_cell_value`.
  if (spill_table_ != nullptr) {
    for (const auto& kv : spill_table_->by_anchor) {
      const SpillRegion& region = kv.second;
      for (std::uint32_t r = 0; r < region.rows; ++r) {
        for (std::uint32_t c = 0; c < region.cols; ++c) {
          if (r == 0U && c == 0U) {
            continue;  // Anchor already occupies a real slot in `rows_`.
          }
          if (cell_at_locked(region.anchor_row + r, region.anchor_col + c) == nullptr) {
            ++total;
          }
        }
      }
    }
  }
  return total;
}

std::vector<CellAddress> Sheet::spill_phantom_addresses() const {
  std::vector<CellAddress> out;
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (spill_table_ == nullptr) {
    return out;
  }
  for (const auto& kv : spill_table_->by_anchor) {
    const SpillRegion& region = kv.second;
    for (std::uint32_t r = 0; r < region.rows; ++r) {
      for (std::uint32_t c = 0; c < region.cols; ++c) {
        if (r == 0U && c == 0U) {
          continue;  // Exclude the anchor; it has a real slot in `rows_`.
        }
        out.push_back(CellAddress{region.anchor_row + r, region.anchor_col + c});
      }
    }
  }
  return out;
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
    if (row < region.anchor_row || col < region.anchor_col ||
        static_cast<std::uint64_t>(row) >= static_cast<std::uint64_t>(region.anchor_row) + region.rows ||
        static_cast<std::uint64_t>(col) >= static_cast<std::uint64_t>(region.anchor_col) + region.cols) {
      continue;
    }
    if (row == region.anchor_row && col == region.anchor_col) {
      return nullptr;
    }
    return &region;
  }
  return nullptr;
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

bool Sheet::commit_spill(std::uint32_t anchor_row, std::uint32_t anchor_col, std::uint32_t rows, std::uint32_t cols,
                         std::vector<Value> cells) {
  // Shape validation. These conditions indicate caller bugs; report them
  // via the debug assert and refuse the registration so release builds
  // remain memory-safe.
  if (rows == 0U || cols == 0U) {
    assert(false && "commit_spill: zero-sized spill region");
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

  // Collision check: scan the footprint excluding the anchor itself.
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c) {
      const std::uint32_t row = anchor_row + r;
      const std::uint32_t col = anchor_col + c;
      if (row == anchor_row && col == anchor_col) {
        continue;
      }
      if (IsCellOccupied(cell_at_locked(row, col))) {
        // Surface #SPILL! at the anchor; preserve the existing literal at
        // the colliding cell.
        std::vector<Cell>& row_cells = rows_[anchor_row];
        Cell& anchor_slot = EnsureSlot(row_cells, anchor_col);
        anchor_slot.cached_value = Value::error(ErrorCode::Spill);
        return false;
      }
      if (spill_region_covering_locked(row, col) != nullptr) {
        std::vector<Cell>& row_cells = rows_[anchor_row];
        Cell& anchor_slot = EnsureSlot(row_cells, anchor_col);
        anchor_slot.cached_value = Value::error(ErrorCode::Spill);
        return false;
      }
    }
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
  std::vector<Cell>& row_cells = rows_[anchor_row];
  Cell& anchor_slot = EnsureSlot(row_cells, anchor_col);
  anchor_slot.cached_value = first_cell;
  return true;
}

namespace {

// Shifts every entry in `metadata` whose anchor lies on or past the
// affected row by the row insert / delete rule encoded in `count` and
// `is_delete`. Entries whose anchor falls inside the deleted interval
// are removed. `Anchor` exposes a mutable `row` field.
template <typename Anchor>
void ShiftRowAnchored(std::vector<Anchor>& items, std::uint32_t row, std::uint32_t count, bool is_delete) {
  std::vector<Anchor> retained;
  retained.reserve(items.size());
  for (Anchor& item : items) {
    if (item.row < row) {
      retained.push_back(std::move(item));
      continue;
    }
    if (is_delete) {
      if (item.row < row + count) {
        continue;  // Anchor inside deleted interval; drop the entry.
      }
      item.row -= count;
    } else {
      // Insert. Anchors at or past `row` shift forward; entries pushed
      // past the sheet bound are dropped.
      const std::uint64_t shifted = static_cast<std::uint64_t>(item.row) + count;
      if (shifted >= Sheet::kMaxRows) {
        continue;
      }
      item.row = static_cast<std::uint32_t>(shifted);
    }
    retained.push_back(std::move(item));
  }
  items = std::move(retained);
}

template <typename Anchor>
void ShiftColAnchored(std::vector<Anchor>& items, std::uint32_t col, std::uint32_t count, bool is_delete) {
  std::vector<Anchor> retained;
  retained.reserve(items.size());
  for (Anchor& item : items) {
    if (item.col < col) {
      retained.push_back(std::move(item));
      continue;
    }
    if (is_delete) {
      if (item.col < col + count) {
        continue;
      }
      item.col -= count;
    } else {
      const std::uint64_t shifted = static_cast<std::uint64_t>(item.col) + count;
      if (shifted >= Sheet::kMaxCols) {
        continue;
      }
      item.col = static_cast<std::uint32_t>(shifted);
    }
    retained.push_back(std::move(item));
  }
  items = std::move(retained);
}

// Rectangular merge / validation range shifter along the row axis.
// Ranges that fall entirely inside the deleted interval are dropped;
// ranges that straddle the deletion are clamped so the surviving rows
// stay contiguous (Excel's "shrink the merge" behaviour). Inserts that
// would push `last_row` past `kMaxRows-1` clamp to the sheet bound.
void ShiftRowRange(MergeRange& range, std::uint32_t row, std::uint32_t count, bool is_delete, bool* out_drop) {
  *out_drop = false;
  if (is_delete) {
    const std::uint32_t del_end = row + count;  // exclusive
    // Both endpoints below the deletion: unchanged.
    if (range.last_row < row) {
      return;
    }
    // Both endpoints inside the deletion: drop the range entirely.
    if (range.first_row >= row && range.last_row < del_end) {
      *out_drop = true;
      return;
    }
    // Split shifts depending on which endpoints fall inside.
    if (range.first_row < row && range.last_row >= row && range.last_row < del_end) {
      // Bottom endpoint inside the deletion; clamp to row-1.
      range.last_row = row - 1U;
      return;
    }
    if (range.first_row >= row && range.first_row < del_end && range.last_row >= del_end) {
      // Top endpoint inside the deletion; clamp to the row after the
      // deletion (which after shift becomes `row`).
      range.first_row = row;
      range.last_row -= count;
      return;
    }
    if (range.first_row < row && range.last_row >= del_end) {
      // Range straddles the entire deletion: shrink by `count`. The
      // top endpoint stays put; the bottom shifts up.
      range.last_row -= count;
      return;
    }
    // Both endpoints past the deletion: shift up.
    range.first_row -= count;
    range.last_row -= count;
    return;
  }
  // Insert. Endpoints at or past `row` shift forward; clamp to bound.
  if (range.last_row < row) {
    return;  // Both endpoints below the insert; unchanged.
  }
  auto shift_one = [count](std::uint32_t value) -> std::uint32_t {
    const std::uint64_t shifted = static_cast<std::uint64_t>(value) + count;
    if (shifted >= Sheet::kMaxRows) {
      return Sheet::kMaxRows - 1U;
    }
    return static_cast<std::uint32_t>(shifted);
  };
  if (range.first_row >= row) {
    range.first_row = shift_one(range.first_row);
  }
  range.last_row = shift_one(range.last_row);
}

void ShiftColRange(MergeRange& range, std::uint32_t col, std::uint32_t count, bool is_delete, bool* out_drop) {
  *out_drop = false;
  if (is_delete) {
    const std::uint32_t del_end = col + count;
    if (range.last_col < col) {
      return;
    }
    if (range.first_col >= col && range.last_col < del_end) {
      *out_drop = true;
      return;
    }
    if (range.first_col < col && range.last_col >= col && range.last_col < del_end) {
      range.last_col = col - 1U;
      return;
    }
    if (range.first_col >= col && range.first_col < del_end && range.last_col >= del_end) {
      range.first_col = col;
      range.last_col -= count;
      return;
    }
    if (range.first_col < col && range.last_col >= del_end) {
      range.last_col -= count;
      return;
    }
    range.first_col -= count;
    range.last_col -= count;
    return;
  }
  if (range.last_col < col) {
    return;
  }
  auto shift_one = [count](std::uint32_t value) -> std::uint32_t {
    const std::uint64_t shifted = static_cast<std::uint64_t>(value) + count;
    if (shifted >= Sheet::kMaxCols) {
      return Sheet::kMaxCols - 1U;
    }
    return static_cast<std::uint32_t>(shifted);
  };
  if (range.first_col >= col) {
    range.first_col = shift_one(range.first_col);
  }
  range.last_col = shift_one(range.last_col);
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

void ShiftSheetMetadata(std::vector<Hyperlink>& hyperlinks, std::vector<CellComment>& comments,
                        std::vector<MergeRange>& merges, std::vector<DataValidation>& validations,
                        std::vector<cf::ConditionalFormat>& conditional_formats, SheetLayout& layout,
                        SheetPrintSettings& print_settings, std::vector<std::unique_ptr<pivot::PivotTable>>& pivots,
                        std::uint32_t index, std::uint32_t count, bool is_delete, bool row_axis) {
  if (row_axis) {
    ShiftRowAnchored(hyperlinks, index, count, is_delete);
    ShiftRowAnchored(comments, index, count, is_delete);
  } else {
    ShiftColAnchored(hyperlinks, index, count, is_delete);
    ShiftColAnchored(comments, index, count, is_delete);
  }
  ShiftRangeList(merges, index, count, is_delete, row_axis);
  for (DataValidation& dv : validations) {
    ShiftRangeList(dv.ranges, index, count, is_delete, row_axis);
  }
  ShiftConditionalFormats(conditional_formats, index, count, is_delete, row_axis);
  if (row_axis) {
    ShiftRowLayouts(layout.row_overrides, index, count, is_delete);
    ShiftBreaks(print_settings.manual_row_breaks, index, count, is_delete, Sheet::kMaxRows);
  } else {
    ShiftColumnLayouts(layout.columns, index, count, is_delete);
    ShiftBreaks(print_settings.manual_col_breaks, index, count, is_delete, Sheet::kMaxCols);
  }
  ShiftPivotAnchors(pivots, index, count, is_delete, row_axis);
}

}  // namespace

void Sheet::insert_rows(std::uint32_t row, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
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
  clear_all_spills();
  ShiftSheetMetadata(hyperlinks_, comments_, merges_, validations_, conditional_formats_, layout_, print_settings_,
                     pivot_tables_, row, count, /*is_delete=*/false, /*row_axis=*/true);
}

void Sheet::delete_rows(std::uint32_t row, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
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
  clear_all_spills();
  ShiftSheetMetadata(hyperlinks_, comments_, merges_, validations_, conditional_formats_, layout_, print_settings_,
                     pivot_tables_, row, count, /*is_delete=*/true, /*row_axis=*/true);
}

void Sheet::insert_cols(std::uint32_t col, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  for (auto& kv : rows_) {
    std::vector<Cell>& cells = kv.second;
    if (col >= cells.size()) {
      continue;
    }
    // Insertion point sits inside this row's populated span. Pad with
    // default Cells starting at `col`. Cells past `kMaxCols` are
    // dropped wholesale.
    const std::size_t old_size = cells.size();
    const std::uint64_t new_size_unbounded = static_cast<std::uint64_t>(old_size) + count;
    const std::size_t new_size = static_cast<std::size_t>(std::min<std::uint64_t>(new_size_unbounded, kMaxCols));
    cells.resize(new_size);
    // Move the tail to its new position. Walk from the back so source
    // and destination never overlap during the swap.
    const std::size_t tail_count = old_size - col;
    for (std::size_t i = 0; i < tail_count; ++i) {
      const std::size_t src = old_size - 1U - i;
      const std::size_t dst = src + count;
      if (dst >= new_size) {
        continue;  // Cell pushed past sheet bound; drop it.
      }
      cells[dst] = std::move(cells[src]);
    }
    // Clear the inserted slots so they read as default-blank.
    const std::size_t clear_end = std::min<std::size_t>(col + count, new_size);
    for (std::size_t i = col; i < clear_end; ++i) {
      cells[i] = Cell{};
    }
  }
  clear_all_spills();
  ShiftSheetMetadata(hyperlinks_, comments_, merges_, validations_, conditional_formats_, layout_, print_settings_,
                     pivot_tables_, col, count, /*is_delete=*/false, /*row_axis=*/false);
}

void Sheet::delete_cols(std::uint32_t col, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  for (auto& kv : rows_) {
    std::vector<Cell>& cells = kv.second;
    if (col >= cells.size()) {
      continue;
    }
    const std::size_t del_end = std::min<std::size_t>(col + count, cells.size());
    cells.erase(cells.begin() + static_cast<std::ptrdiff_t>(col), cells.begin() + static_cast<std::ptrdiff_t>(del_end));
  }
  clear_all_spills();
  ShiftSheetMetadata(hyperlinks_, comments_, merges_, validations_, conditional_formats_, layout_, print_settings_,
                     pivot_tables_, col, count, /*is_delete=*/true, /*row_axis=*/false);
}

void Sheet::clear_spill(std::uint32_t anchor_row, std::uint32_t anchor_col) noexcept {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  clear_spill_locked(anchor_row, anchor_col);
}

void Sheet::clear_all_spills() noexcept {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  while (spill_table_ != nullptr && !spill_table_->by_anchor.empty()) {
    const CellAddress anchor = spill_table_->by_anchor.begin()->first;
    clear_spill_locked(anchor.row, anchor.col);
  }
}

void Sheet::clear_spill_locked(std::uint32_t anchor_row, std::uint32_t anchor_col) noexcept {
  if (spill_table_ == nullptr) {
    return;
  }
  const CellAddress anchor_addr{anchor_row, anchor_col};
  const auto it = spill_table_->by_anchor.find(anchor_addr);
  if (it == spill_table_->by_anchor.end()) {
    return;
  }
  spill_table_->by_anchor.erase(it);
}

}  // namespace formulon
