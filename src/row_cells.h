//
// Per-row sparse cell storage.

#ifndef FORMULON_ROW_CELLS_H_
#define FORMULON_ROW_CELLS_H_

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <iterator>
#include <limits>
#include <memory>
#include <optional>
#include <utility>
#include <vector>

#include "cell.h"
#include "utils/a1_ref.h"

namespace formulon {

/// Hash for `CellAddress` suitable for `std::unordered_map`.
///
/// Excel addresses cap at row < 2^21 and col < 2^14 so a simple
/// `(row * 31) + col` mix is collision-free in the usable range and faster
/// than a generic 64-bit splat. The function is `noexcept` because the
/// underlying field accesses cannot throw.
struct CellAddressHash {
  std::size_t operator()(CellAddress a) const noexcept {
    std::size_t h = static_cast<std::size_t>(a.row);
    h = h * 31U + static_cast<std::size_t>(a.col);
    return h;
  }
};

/// Blank columns `[first, last]` (inclusive) of one row that carry their own
/// format `xf` without a stored cell.
struct StyleRun {
  std::uint16_t first = 0;
  std::uint16_t last = 0;
  std::uint32_t xf = 0;
};

/// A coordinate's own format: the stored cell's xf and explicit flag, or a
/// run's xf (always explicit).
struct StyleSlot {
  std::uint32_t xf = 0;
  bool explicit_xf = false;
};

/// One `from -> to` pair of a remap; a remap's pairs are sorted by `from`
/// with no repeats.
struct XfRemap {
  std::uint32_t from = 0;
  std::uint32_t to = 0;
};

/// An xf mapping in table form (`pairs`, sorted by `from` with no repeats;
/// an xf with no pair maps to itself) or constant form (every xf maps to
/// `constant`).
struct XfMap {
  const XfRemap* pairs = nullptr;
  std::size_t count = 0;
  std::uint32_t constant = 0;
  bool is_constant = false;

  std::uint32_t operator()(std::uint32_t xf) const noexcept;
};

/// One populated row's cells: a sorted array of column keys and a parallel
/// array of cells, so a row's memory follows the number of columns written
/// to it, not the distance between its extreme columns.
///
/// Only written columns are stored; a column between two written ones is
/// absent and `find()` reports it as `nullptr`. Iteration is const and yields
/// `Entry{col, cell}` in ascending column order; mutation goes through
/// `find`, `ensure` and the structural edits.
///
/// Blank columns with their own format are held as style runs beside the
/// stored cells, never as cells. No column is both stored and in a run, and
/// runs are sorted, disjoint, with touching equal-xf runs merged. The stored
/// cell accessors (`empty`, iteration, `entries_in`, `count_in`, `find`)
/// see stored cells only; the structural edits move both together.
///
/// Every column range argument is half-open, `[first, end)`, with a 64-bit
/// `end` so `last_col + 1` never wraps.
class RowCells {
 public:
  /// One stored column and its cell.
  struct Entry {
    std::uint32_t col;
    const Cell& cell;
  };

  /// Forward iterator over stored columns; dereferences to `Entry` by value.
  class const_iterator {
   public:
    using iterator_category = std::forward_iterator_tag;
    using value_type = Entry;
    using difference_type = std::ptrdiff_t;
    using pointer = void;
    using reference = Entry;

    const_iterator() noexcept = default;
    Entry operator*() const noexcept { return Entry{*col_, *cell_}; }
    const_iterator& operator++() noexcept {
      ++col_;
      ++cell_;
      return *this;
    }
    const_iterator operator++(int) noexcept {
      const_iterator prev = *this;
      ++*this;
      return prev;
    }
    bool operator==(const const_iterator& other) const noexcept { return col_ == other.col_; }
    bool operator!=(const const_iterator& other) const noexcept { return col_ != other.col_; }

   private:
    friend class RowCells;
    const_iterator(const std::uint16_t* col, const Cell* cell) noexcept : col_(col), cell_(cell) {}

    const std::uint16_t* col_ = nullptr;
    const Cell* cell_ = nullptr;
  };

  /// Lightweight view over a contiguous slice of a row's entries.
  class EntryRange {
   public:
    const_iterator begin() const noexcept { return begin_; }
    const_iterator end() const noexcept { return end_; }
    bool empty() const noexcept { return begin_ == end_; }

   private:
    friend class RowCells;
    EntryRange(const_iterator b, const_iterator e) noexcept : begin_(b), end_(e) {}

    const_iterator begin_;
    const_iterator end_;
  };

  /// True when no cell is stored; runs are not looked at.
  bool empty() const noexcept { return cols_.empty(); }

  /// True when a cell is stored or a run exists.
  bool has_any() const noexcept { return !cols_.empty() || runs_ != nullptr; }

  /// Number of columns held.
  std::size_t stored_count() const noexcept { return cols_.size(); }

  /// The cell stored at `col`, or `nullptr` when that column was never written.
  const Cell* find(std::uint32_t col) const noexcept {
    const std::size_t index = index_of(col);
    return index == kNpos ? nullptr : &cells_[index];
  }
  Cell* find(std::uint32_t col) noexcept {
    const std::size_t index = index_of(col);
    return index == kNpos ? nullptr : &cells_[index];
  }

  /// Returns the cell at `col`, inserting a default one when absent. Inserting
  /// may move the cells to its right; their `cached_text_owned` payload is
  /// heap-stable across the move. A column inside a run leaves the run, and
  /// the new cell takes the run's xf as an explicit format.
  Cell& ensure(std::uint32_t col);

  /// Every run, ascending (an empty vector when there is none).
  const std::vector<StyleRun>& runs() const noexcept;

  /// The runs intersecting `[first, end)`, unclipped, as a pointer range.
  std::pair<const StyleRun*, const StyleRun*> runs_in(std::uint32_t first, std::uint64_t end) const noexcept;

  /// Number of run columns inside `[first, end)`.
  std::uint32_t run_count_in(std::uint32_t first, std::uint64_t end) const noexcept;

  /// True when a stored or run column lies inside `[first, end)`.
  bool any_in(std::uint32_t first, std::uint64_t end) const noexcept {
    return count_in(first, end) != 0U || run_count_in(first, end) != 0U;
  }

  /// The own format at `col`: a stored cell's `{xf_index, has_explicit_xf}`
  /// (a stored cell owns its style), else a run's `{xf, true}`, else nullopt.
  std::optional<StyleSlot> find_style(std::uint32_t col) const noexcept;

  /// Covers the columns of `[first, end)` that hold no stored cell with runs
  /// of `xf`, replacing the runs there. Returns how many runs the row gained.
  std::size_t set_runs(std::uint32_t first, std::uint64_t end, std::uint32_t xf);

  /// Removes run coverage from `[first, end)`.
  void clear_runs(std::uint32_t first, std::uint64_t end);

  /// Gives every run column inside `[first, end)` the xf `map(xf)`. Returns
  /// how many runs the row gained.
  std::size_t remap_runs(std::uint32_t first, std::uint64_t end, const XfMap& map);

  /// Erases the cell stored at `col` when it holds no content (no formula,
  /// cached value, phonetic state or array role). Returns its xf when it
  /// owned a format (`xf_index != 0 || has_explicit_xf`), for the caller to
  /// write back as a run; nullopt otherwise, and for a column with content
  /// or nothing stored. Invalidates references to this and later cells.
  std::optional<std::uint32_t> demote(std::uint32_t col);

  const_iterator begin() const noexcept { return const_iterator(cols_.data(), cells_.data()); }
  const_iterator end() const noexcept {
    return const_iterator(cols_.data() + cols_.size(), cells_.data() + cells_.size());
  }

  /// Stored entries whose column lies in `[first, end)`.
  EntryRange entries_in(std::uint32_t first, std::uint64_t end) const noexcept {
    const std::size_t lo = lower_index(first);
    const std::size_t hi = std::max(lo, lower_index(end));
    return EntryRange(const_iterator(cols_.data() + lo, cells_.data() + lo),
                      const_iterator(cols_.data() + hi, cells_.data() + hi));
  }

  /// Stored entries whose column is `>= first`.
  EntryRange entries_from(std::uint32_t first) const noexcept {
    const std::size_t lo = lower_index(first);
    return EntryRange(const_iterator(cols_.data() + lo, cells_.data() + lo), end());
  }

  /// Number of stored columns in `[first, end)`.
  std::size_t count_in(std::uint32_t first, std::uint64_t end) const noexcept {
    const std::size_t lo = lower_index(first);
    const std::size_t hi = lower_index(end);
    return hi > lo ? hi - lo : 0U;
  }

  /// Adds `count` to every key and run column at or after `col` (a run
  /// straddling `col` splits there); keys pushed to `kMaxCols` or beyond are
  /// dropped with their cells, a run pushed partly past is truncated.
  void shift_cols_from(std::uint32_t col, std::uint32_t count);

  /// Drops the cell stored at `col` and the run column there, leaving the
  /// other columns where they are.
  void erase(std::uint32_t col);

  /// Drops the keys and run columns in `[col, col + count)` and subtracts
  /// `count` from every column after the band.
  void erase_cols(std::uint32_t col, std::uint32_t count);

  /// Drops the keys and run columns in `[first, end)`; other columns stay.
  void erase_in(std::uint32_t first, std::uint64_t end);

  /// Moves the cells and run columns in `[first, end)` to the same columns of
  /// `dest` (another row), replacing what `dest` held there.
  void move_in_to(std::uint32_t first, std::uint64_t end, RowCells& dest);

  /// Heap bytes reserved by the key, cell and run arrays.
  std::size_t capacity_bytes() const noexcept {
    std::size_t bytes = cols_.capacity() * sizeof(std::uint16_t) + cells_.capacity() * sizeof(Cell);
    if (runs_ != nullptr) {
      bytes += sizeof(std::vector<StyleRun>) + runs_->capacity() * sizeof(StyleRun);
    }
    return bytes;
  }

 private:
  static constexpr std::size_t kNpos = static_cast<std::size_t>(-1);

  /// Inserts a default cell at `col`, which holds none.
  Cell& insert_cell(std::uint32_t col);

  /// Index of the first run whose `last >= col` (`runs_` non-null).
  std::size_t run_lower(std::uint64_t col) const noexcept;

  /// Splits the run straddling `col` so a run boundary falls there; returns
  /// the index of the first run starting at or after `col` (`runs_` non-null).
  std::size_t split_runs_at(std::uint64_t col);

  /// Merges touching equal-xf runs and releases an empty vector.
  void normalize_runs();

  /// Removes run column `col` and returns its xf, or nullopt when no run
  /// holds it.
  std::optional<std::uint32_t> take_run_column(std::uint32_t col);

  /// Number of runs held.
  std::size_t run_total() const noexcept { return runs_ == nullptr ? 0U : runs_->size(); }

  /// Checks that no column is both stored and run, and that runs are sorted,
  /// disjoint and merged. For assertions.
  bool invariants_hold() const noexcept;

  /// Index of `col` in `cols_`, or `kNpos`. Tries the dense-prefix guess first.
  std::size_t index_of(std::uint32_t col) const noexcept {
    if (cols_.empty() || col < cols_.front()) {
      return kNpos;
    }
    const std::size_t guess = std::min<std::size_t>(col - cols_.front(), cols_.size() - 1U);
    if (cols_[guess] == col) {
      return guess;
    }
    const auto it = std::lower_bound(cols_.begin(), cols_.begin() + static_cast<std::ptrdiff_t>(guess), col);
    return (it != cols_.begin() + static_cast<std::ptrdiff_t>(guess) && *it == col)
               ? static_cast<std::size_t>(it - cols_.begin())
               : kNpos;
  }

  /// Index of the first key `>= col`. Keys are unique and ascending, so
  /// `cols_[i] >= front + i` and the answer never exceeds `col - front`; a hit
  /// there proves the prefix is dense.
  std::size_t lower_index(std::uint64_t col) const noexcept {
    if (cols_.empty() || col <= cols_.front()) {
      return 0U;
    }
    if (col > cols_.back()) {
      return cols_.size();
    }
    const std::size_t guess = static_cast<std::size_t>(col - cols_.front());
    if (guess < cols_.size() && cols_[guess] == col) {
      return guess;
    }
    const auto last = cols_.begin() + static_cast<std::ptrdiff_t>(std::min(guess, cols_.size()));
    return static_cast<std::size_t>(std::lower_bound(cols_.begin(), last, col,
                                                     [](std::uint16_t key, std::uint64_t value) {
                                                       return static_cast<std::uint64_t>(key) < value;
                                                     }) -
                                    cols_.begin());
  }

  std::vector<std::uint16_t> cols_;  // ascending, unique
  static_assert(a1::kMaxCols - 1U <= std::numeric_limits<std::uint16_t>::max(), "column keys must fit uint16_t");
  std::vector<Cell> cells_;                      // cells_[i] is column cols_[i]
  std::unique_ptr<std::vector<StyleRun>> runs_;  // null when the row has no run
};

}  // namespace formulon

#endif  // FORMULON_ROW_CELLS_H_
