// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
// The table is two parallel maps:
//
//   * `by_anchor`: anchor cell -> SpillRegion (owns the cell payload).
//   * `covering` : phantom cell -> anchor cell. Anchor cells themselves are
//                  *not* present in this map; lookups for an anchor go
//                  through `by_anchor` directly.
//
// Both maps use `CellAddressHash`. Iteration order is undefined; consumers
// that need a deterministic order must sort externally.
struct SpillTable {
  std::unordered_map<CellAddress, SpillRegion, CellAddressHash> by_anchor;
  std::unordered_map<CellAddress, CellAddress, CellAddressHash> covering;
};

// ---------------------------------------------------------------------------
// Special members (must be defined here where SpillTable is complete).
// ---------------------------------------------------------------------------

Sheet::Sheet(std::string name) : name_(std::move(name)) {}
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

// Returns true when `(row, col)` is "occupied" for the purposes of a spill
// collision check: a non-default cell (literal value or formula) lives there.
// The anchor cell of the would-be spill is excluded from this check by the
// caller.
bool IsCellOccupied(const Sheet& sheet, std::uint32_t row, std::uint32_t col) noexcept {
  const Cell* c = sheet.cell_at(row, col);
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

  // Eager invalidation: writing to a phantom mutates the spilled area, so
  // the spill must be dropped. The anchor's stored `cached_value` is left
  // untouched by `clear_spill`; the next evaluation pass will recompute it
  // (and either re-spill or surface `#SPILL!`).
  if (const SpillRegion* covering = spill_region_covering(row, col); covering != nullptr) {
    clear_spill(covering->anchor_row, covering->anchor_col);
  }

  std::vector<Cell>& row_cells = rows_[row];
  Cell& slot = EnsureSlot(row_cells, col);
  slot.formula_text.clear();
  slot.cached_value = v;
}

void Sheet::set_cell_formula(std::uint32_t row, std::uint32_t col, std::string formula) {
  assert(row < kMaxRows && col < kMaxCols);

  // Same eager invalidation rationale as `set_cell_value`.
  if (const SpillRegion* covering = spill_region_covering(row, col); covering != nullptr) {
    clear_spill(covering->anchor_row, covering->anchor_col);
  }

  std::vector<Cell>& row_cells = rows_[row];
  Cell& slot = EnsureSlot(row_cells, col);
  slot.formula_text = std::move(formula);
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
  std::vector<Cell>& row_cells = rows_[row];
  Cell& slot = EnsureSlot(row_cells, col);
  slot.phonetic_text.assign(phonetic.begin(), phonetic.end());
}

void Sheet::set_cell_xf_index(std::uint32_t row, std::uint32_t col, std::uint32_t xf_index) {
  assert(row < kMaxRows && col < kMaxCols);
  const auto it = rows_.find(row);
  if (it == rows_.end()) {
    return;
  }
  std::vector<Cell>& row_cells = it->second;
  if (col >= row_cells.size()) {
    return;
  }
  row_cells[col].xf_index = xf_index;
}

const Cell* Sheet::cell_at(std::uint32_t row, std::uint32_t col) const noexcept {
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
  std::size_t total = 0;
  for (const auto& kv : rows_) {
    total += kv.second.size();
  }
  return total;
}

// ---------------------------------------------------------------------------
// Spill API
// ---------------------------------------------------------------------------

const SpillRegion* Sheet::spill_region_at_anchor(std::uint32_t row, std::uint32_t col) const noexcept {
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
  if (spill_table_ == nullptr) {
    return nullptr;
  }
  const auto it = spill_table_->covering.find(CellAddress{row, col});
  if (it == spill_table_->covering.end()) {
    return nullptr;
  }
  const auto anchor_it = spill_table_->by_anchor.find(it->second);
  if (anchor_it == spill_table_->by_anchor.end()) {
    // Defensive: the reverse map should never reference a missing anchor.
    return nullptr;
  }
  return &anchor_it->second;
}

Value Sheet::resolve_cell_value(std::uint32_t row, std::uint32_t col) const noexcept {
  if (const SpillRegion* covering = spill_region_covering(row, col); covering != nullptr) {
    const std::uint32_t r_off = row - covering->anchor_row;
    const std::uint32_t c_off = col - covering->anchor_col;
    const std::size_t index =
        static_cast<std::size_t>(r_off) * static_cast<std::size_t>(covering->cols) + static_cast<std::size_t>(c_off);
    return covering->cells[index];
  }
  if (const Cell* c = cell_at(row, col); c != nullptr) {
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

  // Drop any region currently anchored at this cell first, regardless of
  // whether the new commit ends up succeeding. The "register over an
  // existing region" case is intentionally idempotent.
  clear_spill(anchor_row, anchor_col);

  // Collision check: scan the footprint excluding the anchor itself.
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c) {
      const std::uint32_t row = anchor_row + r;
      const std::uint32_t col = anchor_col + c;
      if (row == anchor_row && col == anchor_col) {
        continue;
      }
      if (IsCellOccupied(*this, row, col)) {
        // Surface #SPILL! at the anchor; preserve the existing literal at
        // the colliding cell.
        std::vector<Cell>& row_cells = rows_[anchor_row];
        Cell& anchor_slot = EnsureSlot(row_cells, anchor_col);
        anchor_slot.cached_value = Value::error(ErrorCode::Spill);
        return false;
      }
      if (spill_region_covering(row, col) != nullptr) {
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

  // Register reverse entries for every phantom (anchor excluded). For a
  // degenerate 1x1 region this loop iterates once and skips, so no entries
  // are written.
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c) {
      const std::uint32_t row = anchor_row + r;
      const std::uint32_t col = anchor_col + c;
      if (row == anchor_row && col == anchor_col) {
        continue;
      }
      spill_table_->covering[CellAddress{row, col}] = anchor_addr;
    }
  }

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

void ShiftSheetMetadata(std::vector<Hyperlink>& hyperlinks, std::vector<CellComment>& comments,
                        std::vector<MergeRange>& merges, std::vector<DataValidation>& validations, std::uint32_t index,
                        std::uint32_t count, bool is_delete, bool row_axis) {
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
  ShiftSheetMetadata(hyperlinks_, comments_, merges_, validations_, row, count, /*is_delete=*/false,
                     /*row_axis=*/true);
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
  ShiftSheetMetadata(hyperlinks_, comments_, merges_, validations_, row, count, /*is_delete=*/true,
                     /*row_axis=*/true);
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
  ShiftSheetMetadata(hyperlinks_, comments_, merges_, validations_, col, count, /*is_delete=*/false,
                     /*row_axis=*/false);
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
  ShiftSheetMetadata(hyperlinks_, comments_, merges_, validations_, col, count, /*is_delete=*/true,
                     /*row_axis=*/false);
}

void Sheet::clear_spill(std::uint32_t anchor_row, std::uint32_t anchor_col) noexcept {
  if (spill_table_ == nullptr) {
    return;
  }
  const CellAddress anchor_addr{anchor_row, anchor_col};
  const auto it = spill_table_->by_anchor.find(anchor_addr);
  if (it == spill_table_->by_anchor.end()) {
    return;
  }
  const SpillRegion& region = it->second;
  for (std::uint32_t r = 0; r < region.rows; ++r) {
    for (std::uint32_t c = 0; c < region.cols; ++c) {
      const std::uint32_t row = region.anchor_row + r;
      const std::uint32_t col = region.anchor_col + c;
      if (row == region.anchor_row && col == region.anchor_col) {
        continue;
      }
      spill_table_->covering.erase(CellAddress{row, col});
    }
  }
  spill_table_->by_anchor.erase(it);
}

}  // namespace formulon
