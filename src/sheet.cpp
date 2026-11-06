//
// Out-of-line implementation of the sparse cell store
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
#include "utils/derived_guid.h"
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
  slot.array_block_cols = 0U;
  slot.array_block_rows = 0U;
  slot.phonetic_runs.clear();
  slot.phonetic_props = PhoneticProperties{};
  slot.phonetic_present = false;
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

/// Copies `src` into the fresh slot `dst`. A Text value viewing the cell's
/// own buffer is re-pointed at the copy's buffer; any other view (a shared
/// string) is kept as is.
void clone_cell(const Cell& src, Cell& dst) {
  dst.formula_text = src.formula_text;
  dst.cached_value = src.cached_value;
  if (src.cached_text_owned != nullptr) {
    dst.cached_text_owned = std::make_unique<std::string>(*src.cached_text_owned);
    if (src.cached_value.is_text() && src.cached_value.as_text().data() == src.cached_text_owned->data()) {
      dst.cached_value = Value::text(*dst.cached_text_owned);
    }
  }
  dst.phonetic_runs = src.phonetic_runs;
  dst.xf_index = src.xf_index;
  dst.phonetic_props = src.phonetic_props;
  dst.dynamic_array = src.dynamic_array;
  dst.has_explicit_xf = src.has_explicit_xf;
  dst.show_phonetic = src.show_phonetic;
  dst.phonetic_present = src.phonetic_present;
  dst.array_block_cols = src.array_block_cols;
  dst.array_block_rows = src.array_block_rows;
}

/// `derived_guid(guid, salt)`, keeping an absent (empty) GUID absent.
void rederive_guid(std::string& guid, std::string_view salt) {
  if (!guid.empty()) {
    guid = derived_guid(guid, salt);
  }
}

/// Removes the `codeName` attribute from the start tag of a raw `<sheetPr>`.
void strip_code_name(std::string& sheet_pr_xml) {
  constexpr std::string_view kAttr = "codeName=";
  const std::size_t tag_end = sheet_pr_xml.find('>');
  for (std::size_t at = sheet_pr_xml.find(kAttr); at != std::string::npos && at < tag_end;
       at = sheet_pr_xml.find(kAttr, at + 1U)) {
    const char before = at > 0U ? sheet_pr_xml[at - 1U] : '\0';
    if (before != ' ' && before != '\t' && before != '\n' && before != '\r') {
      continue;
    }
    const std::size_t value = at + kAttr.size();
    if (value >= sheet_pr_xml.size() || (sheet_pr_xml[value] != '"' && sheet_pr_xml[value] != '\'')) {
      continue;
    }
    const std::size_t close = sheet_pr_xml.find(sheet_pr_xml[value], value + 1U);
    if (close == std::string::npos || close > tag_end) {
      return;
    }
    sheet_pr_xml.erase(at - 1U, close + 1U - (at - 1U));
    return;
  }
}

}  // namespace

// ---------------------------------------------------------------------------
// Special members (must be defined here where SpillTable is complete).
// ---------------------------------------------------------------------------

Sheet::Sheet(std::string name) : name_(std::move(name)) {}

// Defaulting these out-of-line keeps the incomplete SpillTable deleter out of
// the header while ensuring new Sheet metadata participates in moves without
// a hand-maintained member list.
Sheet::Sheet(Sheet&&) noexcept = default;
Sheet& Sheet::operator=(Sheet&&) noexcept = default;

Sheet::~Sheet() = default;

namespace sheet_copy {

SpillTablePtr::SpillTablePtr(const SpillTablePtr&) noexcept : std::unique_ptr<SpillTable>() {}

RowStore::RowStore(const RowStore& other) : std::unordered_map<std::uint32_t, RowCells>() {
  reserve(other.size());
  for (const auto& [row, cells] : other) {
    RowCells& dst = (*this)[row];
    for (const auto& [col, cell] : cells) {
      clone_cell(cell, dst.ensure(col));
    }
  }
}

PivotList::PivotList(const PivotList& other) : std::vector<std::unique_ptr<pivot::PivotTable>>() {
  reserve(other.size());
  for (const std::unique_ptr<pivot::PivotTable>& table : other) {
    push_back(table->clone());
  }
}

}  // namespace sheet_copy

Sheet::Sheet(const Sheet&) = default;

Sheet Sheet::clone_as(std::string name) const {
  Sheet copy = [this] {
    const std::lock_guard<std::mutex> guard(*spill_mutex_);
    return Sheet(*this);
  }();
  copy.name_ = std::move(name);
  const std::string& salt = copy.name_;
  // These name the source's package identity and parts.
  copy.source_sheet_id_ = 0;
  copy.drawing_rel_target_.clear();
  copy.next_drawing_id_ = 2;
  copy.comment_vml_path_.clear();

  for (Hyperlink& link : copy.hyperlinks_) {
    link.rid.clear();
  }
  for (CellComment& comment : copy.comments_) {
    rederive_guid(comment.uid, salt);
  }
  // `derived_guid` is a pure function, so a reply's `parent_id` lands on its
  // thread's new id without a lookup table.
  for (ThreadedComment& comment : copy.threaded_comments_) {
    rederive_guid(comment.id, salt);
    rederive_guid(comment.parent_id, salt);
    for (Mention& mention : comment.mentions) {
      rederive_guid(mention.mention_id, salt);
    }
  }
  for (DataValidation& validation : copy.validations_) {
    rederive_attr_guid(validation.extra_attrs, "xr:uid", salt);
  }
  copy.view_.tab_selected = false;
  copy.print_settings_.printer_settings_rid.clear();
  strip_code_name(copy.print_settings_.sheet_pr_xml);
  return copy;
}

void Sheet::add_pivot_table(std::unique_ptr<pivot::PivotTable> table) {
  if (table == nullptr) {
    return;
  }
  pivot_tables_.push_back(std::move(table));
}

namespace {

// (row << 32 | col): packed keys sort in the same order as the pairs.
std::uint64_t packed_cell_key(std::uint32_t row, std::uint32_t col) noexcept {
  return (static_cast<std::uint64_t>(row) << 32U) | col;
}

// Indices of `keys` in ascending key order.
std::vector<std::uint32_t> sorted_key_order(const std::vector<std::uint64_t>& keys) {
  std::vector<std::uint32_t> order;
  sorted_index_order_ascending(order, keys);
  return order;
}

}  // namespace

std::vector<std::pair<std::uint32_t, std::uint32_t>> Sheet::comment_anchor_set() const {
  std::vector<std::uint64_t> keys;
  keys.reserve(comments_.size() + threaded_comments_.size());
  for (const CellComment& c : comments_) {
    keys.push_back(packed_cell_key(c.row, c.col));
  }
  for (const ThreadedComment& c : threaded_comments_) {
    keys.push_back(packed_cell_key(c.row, c.col));
  }
  const std::vector<std::uint32_t> order = sorted_key_order(keys);
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

std::vector<CommentShape> Sheet::comment_shapes() const {
  std::vector<CommentShape> shapes;
  const auto anchors = comment_anchor_set();
  shapes.reserve(anchors.size());
  // Comments ordered by packed (row, col); equal keys keep their input order,
  // so the lower bound of an anchor is its first comment.
  std::vector<std::uint64_t> keys;
  keys.reserve(comments_.size());
  for (const CellComment& c : comments_) {
    keys.push_back(packed_cell_key(c.row, c.col));
  }
  const std::vector<std::uint32_t> order = sorted_key_order(keys);
  for (const auto& [row, col] : anchors) {
    CommentShape shape;
    shape.row = row;
    shape.col = col;
    const std::uint64_t key = packed_cell_key(row, col);
    const auto at = std::lower_bound(order.begin(), order.end(), key,
                                     [&keys](std::uint32_t index, std::uint64_t k) { return keys[index] < k; });
    if (at != order.end() && keys[*at] == key) {
      const CellComment& c = comments_[*at];
      shape.visible = c.visible;
      shape.has_box = c.has_box;
      shape.box = c.box;
      shape.xml = c.vml_shape;
    }
    shapes.push_back(std::move(shape));
  }
  return shapes;
}

Cell& Sheet::ensure_overwritten_cell_locked(std::uint32_t row, std::uint32_t col) {
  // A write can replace either a spill anchor or a phantom; the whole region
  // must disappear before storage is updated.
  if (const std::optional<CellAddress> anchor = spill_displaced_by_write(spill_table_.get(), row, col)) {
    clear_spill_locked(anchor->row, anchor->col);
  }
  return rows_[row].ensure(col);
}

void Sheet::set_cell_value(std::uint32_t row, std::uint32_t col, Value v) {
  // Bounds checks are advisory: callers above this layer (parser, OOXML
  // reader) own coordinate validation. A debug assert catches programming
  // errors without imposing a release-mode branch.
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);

  Cell& slot = ensure_overwritten_cell_locked(row, col);
  reset_to_literal(slot);
  slot.cached_value = v;
  if (!v.is_text()) {
    slot.cached_text_owned.reset();
  }
  index_formula_cell_locked(row, col, false);
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_text(std::uint32_t row, std::uint32_t col, std::string_view text) {
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  Cell& slot = ensure_overwritten_cell_locked(row, col);
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

  Cell& slot = ensure_overwritten_cell_locked(row, col);
  slot.formula_text = std::move(formula);
  slot.phonetic_runs.clear();
  slot.phonetic_props = PhoneticProperties{};
  slot.phonetic_present = false;
  slot.cached_value = Value::blank();
  slot.cached_text_owned.reset();
  index_formula_cell_locked(row, col, !slot.formula_text.empty());
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_dynamic_array(std::uint32_t row, std::uint32_t col, bool dynamic) {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (Cell* cell = find_formula_cell(rows_, row, col); cell != nullptr) {
    cell->dynamic_array = dynamic;
  }
}

void Sheet::set_cell_array_block(std::uint32_t row, std::uint32_t col, std::uint32_t rows, std::uint32_t cols) {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (Cell* cell = find_formula_cell(rows_, row, col); cell != nullptr) {
    cell->array_block_rows = cols == 0U ? 0U : rows;
    cell->array_block_cols = static_cast<std::uint16_t>(rows == 0U ? 0U : cols);
  }
}

void Sheet::set_cell_formula_text(std::uint32_t row, std::uint32_t col, std::string formula) {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  if (Cell* cell = find_formula_cell(rows_, row, col); cell != nullptr && !formula.empty()) {
    cell->formula_text = std::move(formula);
  }
}

void Sheet::set_cell_cached_value_locked(Cell& slot, Value v) {
  if (v.is_text()) {
    // Deep-copy the Text payload into a heap-stable `std::string` owned by
    // the Cell so the stored `cached_value` no longer references whatever
    // buffer the caller used (the recalc engine's per-evaluation `Arena`, or a
    // spill region's strings about to be replaced). Allocating a fresh `unique_ptr<std::string>` per
    // write keeps the bytes at a fixed heap address that survives the Cell
    // being relocated by a row's insert or growth — a bare `std::string`
    // would relocate its inline (SSO) bytes on every Cell move and dangle the
    // `string_view` we are about to store.
    //
    // Construct from `v.as_text()` directly: even when the caller passes the
    // cell's own current cached text back through this API, the new string
    // allocates its own buffer before we replace `cached_text_owned`, so the
    // source bytes remain valid for the duration of the construction.
    auto owned = std::make_unique<std::string>(v.as_text());
    const std::string_view view(*owned);
    slot.cached_text_owned = std::move(owned);
    slot.cached_value = Value::text(view);
  } else {
    slot.cached_value = v;
    slot.cached_text_owned.reset();
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
  set_cell_cached_value_locked(slot, std::move(v));
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
    // the surface text end to end. The span is computed now rather than
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

void Sheet::set_cell_phonetic_props(std::uint32_t row, std::uint32_t col, PhoneticProperties props, bool explicit_pr) {
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  RowCells& row_cells = rows_[row];
  Cell& slot = row_cells.ensure(col);
  slot.phonetic_props = props;
  slot.phonetic_present = explicit_pr;
  cell_enumeration_revision_.bump();
}

void Sheet::set_cell_show_phonetic(std::uint32_t row, std::uint32_t col, bool show) {
  assert(row < kMaxRows && col < kMaxCols);

  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  RowCells& row_cells = rows_[row];
  row_cells.ensure(col).show_phonetic = show;
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

void Sheet::clear_cell_xf_index(std::uint32_t row, std::uint32_t col) {
  assert(row < kMaxRows && col < kMaxCols);
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  const auto it = rows_.find(row);
  if (it == rows_.end()) {
    return;
  }
  Cell* slot = it->second.find(col);
  if (slot == nullptr) {
    return;
  }
  slot->xf_index = 0U;
  slot->has_explicit_xf = false;
  cell_enumeration_revision_.bump();
}

void Sheet::restore_cell_style(std::uint32_t row, std::uint32_t col, std::uint32_t xf_index, bool explicit_xf) {
  if (!explicit_xf && xf_index == 0U) {
    clear_cell_xf_index(row, col);
    return;
  }
  assert(row < kMaxRows && col < kMaxCols);
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  Cell& slot = rows_[row].ensure(col);
  slot.xf_index = xf_index;
  slot.has_explicit_xf = explicit_xf;
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

void Sheet::erase_cell(std::uint32_t row, std::uint32_t col) {
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  const auto it = rows_.find(row);
  if (it == rows_.end()) {
    return;
  }
  it->second.erase(col);
  index_formula_cell_locked(row, col, false);
  cell_enumeration_revision_.bump();
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
  // phantom that coincides with a stored cell is already counted above, so
  // only phantoms absent from `rows_` add to the total. This keeps the count
  // aligned with the flat enumeration exposed through the C ABI, which
  // surfaces phantoms via `resolve_cell_value`.
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
    for (const auto& [col, cell] : cells) {
      if (!cell.formula_text.empty()) {
        formula_cells_.insert(formula_cell_key(row, col));
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
  auto write = merges_.begin();
  for (auto read = merges_.begin(); read != merges_.end(); ++read) {
    if (!intersects(*read, range)) {
      if (write != read) {
        *write = std::move(*read);
      }
      ++write;
      continue;
    }
    removed.push_back(*read);
  }
  merges_.erase(write, merges_.end());
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
  out.array_block_rows_ = 0U;
  out.array_block_cols_ = 0U;
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
    out.array_block_rows_ = cell->array_block_rows;
    out.array_block_cols_ = cell->array_block_cols;
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
