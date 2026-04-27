// Copyright 2026 libraz. Licensed under the MIT License.
//
// Out-of-line implementation of the row-sparse, column-dense cell store
// owned by `Sheet` and the heap-owned spill-region table. See `sheet.h` for
// the storage-layer and spill API contracts.

#include "sheet.h"

#include <cassert>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include "cell.h"
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
