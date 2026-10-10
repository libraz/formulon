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
#include "phonetic.h"
#include "pivot/pivot_table.h"
#include "sheet_spill_table.h"
#include "utils/arena.h"
#include "utils/index_sort.h"
#include "utils/utf8_length.h"
#include "value.h"

namespace formulon {

namespace {

/// Anchor of the committed or blocked spill that a cell write at `(row, col)`
/// displaces: the cell's own anchor entry, else the region covering it.
std::optional<CellAddress> spill_displaced_by_write(const SpillTable* table, std::uint32_t row, std::uint32_t col) {
  if (table == nullptr) {
    return std::nullopt;
  }
  const CellAddress address{row, col};
  if (table->by_anchor.find(address) != table->by_anchor.end() ||
      table->blocked_by_anchor.find(address) != table->blocked_by_anchor.end()) {
    return address;
  }
  if (const SpillRegion* covering = find_region_covering(*table, row, col); covering != nullptr) {
    return CellAddress{covering->anchor_row, covering->anchor_col};
  }
  return std::nullopt;
}

/// Turns `slot` into a literal: drops the formula and the annotations the
/// previous value carried.
void reset_to_literal(Cell& slot) {
  slot.formula_text.clear();
  slot.dynamic_array = false;
  slot.phonetic_runs.clear();
  slot.phonetic_props = PhoneticProperties{};
}

/// The stored cell at `(row, col)` when it holds a formula, else null.
Cell* find_formula_cell(std::unordered_map<std::uint32_t, RowCells>& rows, std::uint32_t row, std::uint32_t col) {
  const auto row_it = rows.find(row);
  if (row_it == rows.end()) {
    return nullptr;
  }
  Cell* cell = row_it->second.find(col);
  return cell != nullptr && !cell->formula_text.empty() ? cell : nullptr;
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

std::vector<std::pair<std::uint32_t, std::uint32_t>> Sheet::comment_anchor_set() const {
  // Packed (row << 32 | col) keys sort in the same order as the pairs.
  std::vector<std::uint64_t> keys;
  keys.reserve(comments_.size() + threaded_comments_.size());
  for (const CellComment& c : comments_) {
    keys.push_back((static_cast<std::uint64_t>(c.row) << 32U) | c.col);
  }
  for (const ThreadedComment& c : threaded_comments_) {
    keys.push_back((static_cast<std::uint64_t>(c.row) << 32U) | c.col);
  }
  std::vector<std::uint32_t> order;
  sorted_index_order(order, static_cast<std::uint32_t>(keys.size()),
                     IndexLess{&keys, [](const void* context, std::uint32_t lhs, std::uint32_t rhs) {
                                 const auto& packed = *static_cast<const std::vector<std::uint64_t>*>(context);
                                 return packed[lhs] < packed[rhs];
                               }});
  std::vector<std::pair<std::uint32_t, std::uint32_t>> anchors;
  anchors.reserve(order.size());
  for (std::size_t i = 0; i < order.size(); ++i) {
    const std::uint64_t k = keys[order[i]];
    if (i == 0 || k != keys[order[i - 1U]]) {
      anchors.emplace_back(static_cast<std::uint32_t>(k >> 32U), static_cast<std::uint32_t>(k));
    }
  }
  return anchors;
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

void Sheet::set_cell_value(std::uint32_t row, std::uint32_t col, Value v) {
  // Bounds checks are advisory: callers above this layer (parser, OOXML
  // reader) own coordinate validation. A debug assert catches programming
  // errors without imposing a release-mode branch.
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);

  // A literal write can replace either the anchor or a phantom. In both
  // cases the complete region must disappear before updating storage.
  if (const std::optional<CellAddress> anchor = spill_displaced_by_write(spill_table_.get(), row, col)) {
    clear_spill_locked(anchor->row, anchor->col);
  }

  RowCells& row_cells = rows_[row];
  Cell& slot = row_cells.ensure(col);
  reset_to_literal(slot);
  slot.cached_value = v;
  index_formula_cell_locked(row, col, false);
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_text(std::uint32_t row, std::uint32_t col, std::string_view text) {
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (const std::optional<CellAddress> anchor = spill_displaced_by_write(spill_table_.get(), row, col)) {
    clear_spill_locked(anchor->row, anchor->col);
  }

  RowCells& row_cells = rows_[row];
  Cell& slot = row_cells.ensure(col);
  reset_to_literal(slot);
  auto owned = std::make_unique<std::string>(text);
  slot.cached_value = Value::text(*owned);
  slot.cached_text_owned = std::move(owned);
  index_formula_cell_locked(row, col, false);
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_formula(std::uint32_t row, std::uint32_t col, std::string formula) {
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);

  // Formula replacement can likewise target an anchor or a phantom.
  if (const std::optional<CellAddress> anchor = spill_displaced_by_write(spill_table_.get(), row, col)) {
    clear_spill_locked(anchor->row, anchor->col);
  }

  RowCells& row_cells = rows_[row];
  Cell& slot = row_cells.ensure(col);
  slot.formula_text = std::move(formula);
  slot.phonetic_runs.clear();
  slot.phonetic_props = PhoneticProperties{};
  slot.cached_value = Value::blank();
  index_formula_cell_locked(row, col, !slot.formula_text.empty());
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_dynamic_array(std::uint32_t row, std::uint32_t col, bool dynamic) {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (Cell* cell = find_formula_cell(rows_, row, col); cell != nullptr) {
    cell->dynamic_array = dynamic;
  }
}

void Sheet::set_cell_formula_text(std::uint32_t row, std::uint32_t col, std::string formula) {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (Cell* cell = find_formula_cell(rows_, row, col); cell != nullptr && !formula.empty()) {
    cell->formula_text = std::move(formula);
  }
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
  slot.phonetic_runs.clear();
  if (!phonetic.empty()) {
    // A caller supplying a bare kana string is annotating the whole cell:
    // there is no span vocabulary on this entry point, so the run covers
    // the surface text end to end. The span is measured now rather than
    // at write time because every value-mutating setter clears the
    // annotation, so the text it describes cannot change underneath it.
    const std::string_view surface = slot.cached_value.is_text() ? slot.cached_value.as_text() : std::string_view{};
    slot.phonetic_runs.push_back(PhoneticRun{0U, utf16_units_in(surface), std::string(phonetic)});
  }
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_phonetic_runs(std::uint32_t row, std::uint32_t col, std::vector<PhoneticRun> runs) {
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  RowCells& row_cells = rows_[row];
  Cell& slot = row_cells.ensure(col);
  slot.phonetic_runs = std::move(runs);
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_phonetic_props(std::uint32_t row, std::uint32_t col, PhoneticProperties props) {
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  RowCells& row_cells = rows_[row];
  Cell& slot = row_cells.ensure(col);
  slot.phonetic_props = props;
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
  slot.has_explicit_xf = true;
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
  const auto scan_row = [&](std::uint32_t row, const RowCells& cells) {
    if (cells.empty()) {
      return;
    }
    const std::size_t begin = std::max<std::size_t>(first_col, cells.first_col());
    const std::size_t end = std::min<std::size_t>(static_cast<std::size_t>(last_col) + 1U, cells.size());
    for (std::size_t col = begin; col < end; ++col) {
      const Cell& cell = cells[col];
      if (!cell.formula_text.empty() || !cell.cached_value.is_blank()) {
        include(row, static_cast<std::uint32_t>(col));
      }
    }
  };
  // Walk whichever of the stored rows and the rectangle's rows is smaller,
  // as `materialised_cells_in_rect_locked` does.
  if (static_cast<std::uint64_t>(rows_.size()) <= static_cast<std::uint64_t>(last_row - first_row) + 1U) {
    for (const auto& [row, cells] : rows_) {
      if (row >= first_row && row <= last_row) {
        scan_row(row, cells);
      }
    }
  } else {
    for (std::uint32_t row = first_row; row <= last_row; ++row) {
      const auto it = rows_.find(row);
      if (it != rows_.end()) {
        scan_row(row, it->second);
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

namespace {

std::uint64_t formula_cell_key(std::uint32_t row, std::uint32_t col) noexcept {
  return (static_cast<std::uint64_t>(col) << 32U) | row;
}

}  // namespace

std::vector<CellAddress> Sheet::formula_cells_in(std::uint32_t first_row, std::uint32_t first_col,
                                                 std::uint32_t last_row, std::uint32_t last_col) const {
  std::vector<CellAddress> out;
  if (!rect_in_grid(first_row, first_col, last_row, last_col)) {
    return out;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  // One ordered interval per column: seek to `(col, first_row)`, take rows up
  // to `last_row`, then seek straight to the next column's `first_row`.
  auto it = formula_cells_.lower_bound(formula_cell_key(first_row, first_col));
  while (it != formula_cells_.end()) {
    const auto col = static_cast<std::uint32_t>(*it >> 32U);
    const auto row = static_cast<std::uint32_t>(*it & 0xFFFFFFFFU);
    if (col > last_col) {
      break;
    }
    if (row < first_row) {
      it = formula_cells_.lower_bound(formula_cell_key(first_row, col));
      continue;
    }
    if (row > last_row) {
      if (col == last_col) {
        break;
      }
      it = formula_cells_.lower_bound(formula_cell_key(first_row, col + 1U));
      continue;
    }
    out.push_back(CellAddress{row, col});
    ++it;
  }
  return out;
}

std::uint64_t Sheet::cells_in_range(std::uint32_t first_row, std::uint32_t first_col, std::uint32_t last_row,
                                    std::uint32_t last_col, std::uint64_t cursor, std::uint32_t limit,
                                    void (*visit)(const RangeCell& cell, void* ctx), void* ctx) const {
  constexpr std::uint64_t kStride = kMaxCols;
  if (!rect_in_grid(first_row, first_col, last_row, last_col) || cursor >= kStride * kMaxRows) {
    return kCellCursorEnd;
  }
  auto start_row = static_cast<std::uint32_t>(cursor / kStride);
  auto start_col = static_cast<std::uint32_t>(cursor % kStride);
  if (start_row < first_row) {
    start_row = first_row;
    start_col = first_col;
  }
  start_col = std::max(start_col, first_col);
  if (start_col > last_col) {
    ++start_row;
    start_col = first_col;
  }
  if (start_row > last_row) {
    return kCellCursorEnd;
  }

  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  std::vector<const SpillRegion*> regions;
  if (spill_table_ != nullptr) {
    for (const auto& [unused, region] : spill_table_->by_anchor) {
      (void)unused;
      if (RectIntersectsSpan(region, start_row, static_cast<std::uint64_t>(last_row) + 1U, first_col,
                             static_cast<std::uint64_t>(last_col) + 1U)) {
        regions.push_back(&region);
      }
    }
  }

  // Phantom-free rectangles taller than the stored row count walk the sorted
  // stored rows instead of probing every row, as `populated_extent` does.
  std::vector<std::uint32_t> stored_rows;
  const bool walk_stored =
      regions.empty() && static_cast<std::uint64_t>(rows_.size()) <= static_cast<std::uint64_t>(last_row - start_row);
  if (walk_stored) {
    for (const auto& entry : rows_) {
      if (entry.first >= start_row && entry.first <= last_row) {
        stored_rows.push_back(entry.first);
      }
    }
    sort_ascending(stored_rows);
  }

  std::uint32_t emitted = 0;
  std::vector<std::uint32_t> cols;
  const auto scan_row = [&](std::uint32_t row) -> std::uint64_t {
    const std::uint32_t col_begin = row == start_row ? start_col : first_col;
    cols.clear();
    const auto row_it = rows_.find(row);
    const RowCells* cells = row_it == rows_.end() ? nullptr : &row_it->second;
    if (cells != nullptr && !cells->empty()) {
      const std::size_t end = std::min<std::size_t>(static_cast<std::size_t>(last_col) + 1U, cells->size());
      for (std::size_t col = std::max<std::size_t>(col_begin, cells->first_col()); col < end; ++col) {
        const Cell& cell = (*cells)[col];
        if (!cell.formula_text.empty() || !cell.cached_value.is_blank()) {
          cols.push_back(static_cast<std::uint32_t>(col));
        }
      }
    }
    bool phantom_cols = false;
    for (const SpillRegion* region : regions) {
      if (row < region->anchor_row || row - region->anchor_row >= region->rows) {
        continue;
      }
      const std::uint32_t span_first = std::max(col_begin, region->anchor_col);
      const std::uint32_t span_last = std::min(last_col, region->anchor_col + region->cols - 1U);
      for (std::uint32_t col = span_first; col <= span_last; ++col) {
        if (row != region->anchor_row || col != region->anchor_col) {
          cols.push_back(col);
          phantom_cols = true;
        }
      }
    }
    if (phantom_cols) {
      sort_ascending(cols);
      cols.erase(std::unique(cols.begin(), cols.end()), cols.end());
    }
    for (const std::uint32_t col : cols) {
      if (emitted == limit) {
        return static_cast<std::uint64_t>(row) * kStride + col;
      }
      RangeCell out;
      out.row = row;
      out.col = col;
      // Same precedence as `read_formula_cell`: formula cell, then phantom,
      // then the stored literal.
      const Cell* cell = cells == nullptr ? nullptr : cells->find(col);
      if (cell != nullptr && !cell->formula_text.empty()) {
        out.formula_text = cell->formula_text;
        out.value = cell->cached_value;
      } else if (const SpillRegion* covering = spill_region_covering_locked(row, col); covering != nullptr) {
        out.value = covering->cells[static_cast<std::size_t>(row - covering->anchor_row) * covering->cols +
                                    (col - covering->anchor_col)];
      } else if (cell != nullptr) {
        out.value = cell->cached_value;
      }
      visit(out, ctx);
      ++emitted;
    }
    return kCellCursorEnd;
  };

  if (walk_stored) {
    for (const std::uint32_t row : stored_rows) {
      if (const std::uint64_t next = scan_row(row); next != kCellCursorEnd) {
        return next;
      }
    }
  } else {
    for (std::uint32_t row = start_row; row <= last_row; ++row) {
      if (const std::uint64_t next = scan_row(row); next != kCellCursorEnd) {
        return next;
      }
    }
  }
  return kCellCursorEnd;
}

void Sheet::index_formula_cell_locked(std::uint32_t row, std::uint32_t col, bool has_formula) {
  if (has_formula) {
    formula_cells_.insert(formula_cell_key(row, col));
  } else {
    formula_cells_.erase(formula_cell_key(row, col));
  }
}

void Sheet::rebuild_formula_index_locked() {
  formula_cells_.clear();
  for (const auto& [row, cells] : rows_) {
    for (std::size_t col = cells.first_col(); col < cells.size(); ++col) {
      if (!cells[col].formula_text.empty()) {
        formula_cells_.insert(formula_cell_key(row, static_cast<std::uint32_t>(col)));
      }
    }
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
        out.push_back(adopt_text_into(text_arena, cell->cached_value));
        continue;
      }
      if (const SpillRegion* region = covering(row, col); region != nullptr) {
        const std::size_t index = static_cast<std::size_t>(row - region->anchor_row) * region->cols +
                                  static_cast<std::size_t>(col - region->anchor_col);
        out.push_back(adopt_text_into(text_arena, region->cells[index]));
        continue;
      }
      out.push_back(cell != nullptr ? adopt_text_into(text_arena, cell->cached_value) : Value::blank());
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
    out_cells.push_back(adopt_text_into(text_arena, value));
  }
  if (out_rows != nullptr) {
    *out_rows = region.rows;
  }
  if (out_cols != nullptr) {
    *out_cols = region.cols;
  }
  return true;
}

}  // namespace formulon
