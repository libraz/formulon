//
// `RowCells` column insertion and removal: the sorted column-key array and its
// parallel cell array move together, and the style runs move with them.

#include "row_cells.h"

#include <algorithm>
#include <cassert>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <utility>
#include <vector>

#include "cell.h"
#include "utils/a1_ref.h"

namespace formulon {

namespace {

/// A stored cell's content (anything beyond its format).
bool HasContent(const Cell& cell) noexcept {
  return !cell.formula_text.empty() || !cell.cached_value.is_blank() || cell.show_phonetic || cell.phonetic_present ||
         !cell.phonetic_runs.empty() || cell.array_block_rows != 0U || cell.dynamic_array;
}

/// `end` clamped to the grid.
std::uint64_t GridStop(std::uint64_t end) noexcept {
  return std::min<std::uint64_t>(end, a1::kMaxCols);
}

// Replaces list[lo, hi) with [first, last) by appending, so no range insert is instantiated.
void splice_runs(std::vector<StyleRun>& list, std::size_t lo, std::size_t hi, const StyleRun* first,
                 const StyleRun* last) {
  std::vector<StyleRun> tail;
  for (std::size_t i = hi; i < list.size(); ++i) {
    tail.push_back(list[i]);
  }
  list.erase(list.begin() + static_cast<std::ptrdiff_t>(lo), list.end());
  for (; first != last; ++first) {
    list.push_back(*first);
  }
  for (const StyleRun& run : tail) {
    list.push_back(run);
  }
}

}  // namespace

std::uint32_t XfMap::operator()(std::uint32_t xf) const noexcept {
  if (is_constant) {
    return constant;
  }
  const XfRemap* stop = pairs + count;
  const XfRemap* it =
      std::lower_bound(pairs, stop, xf, [](const XfRemap& pair, std::uint32_t value) { return pair.from < value; });
  return it != stop && it->from == xf ? it->to : xf;
}

Cell& RowCells::insert_cell(std::uint32_t col) {
  // Right-append fast path: rows are mostly written left to right.
  if (cols_.empty() || col > cols_.back()) {
    cols_.push_back(static_cast<std::uint16_t>(col));
    cells_.emplace_back();
    return cells_.back();
  }
  const std::size_t index = lower_index(col);
  cols_.insert(cols_.begin() + static_cast<std::ptrdiff_t>(index), static_cast<std::uint16_t>(col));
  return *cells_.emplace(cells_.begin() + static_cast<std::ptrdiff_t>(index));
}

Cell& RowCells::ensure(std::uint32_t col) {
  assert(col < a1::kMaxCols);
  if (const std::optional<std::uint32_t> run_xf = take_run_column(col)) {
    Cell& cell = insert_cell(col);
    cell.xf_index = *run_xf;
    cell.has_explicit_xf = true;
    assert(invariants_hold());
    return cell;
  }
  if (Cell* existing = find(col); existing != nullptr) {
    return *existing;
  }
  return insert_cell(col);
}

const std::vector<StyleRun>& RowCells::runs() const noexcept {
  static const std::vector<StyleRun> kNoRuns;
  return runs_ == nullptr ? kNoRuns : *runs_;
}

std::pair<const StyleRun*, const StyleRun*> RowCells::runs_in(std::uint32_t first, std::uint64_t end) const noexcept {
  if (runs_ == nullptr || end <= first) {
    return {nullptr, nullptr};
  }
  const StyleRun* data = runs_->data();
  const StyleRun* lo = data + run_lower(first);
  const StyleRun* hi = std::lower_bound(lo, data + runs_->size(), end, [](const StyleRun& run, std::uint64_t value) {
    return static_cast<std::uint64_t>(run.first) < value;
  });
  return {lo, hi};
}

std::uint32_t RowCells::run_count_in(std::uint32_t first, std::uint64_t end) const noexcept {
  const auto [lo, hi] = runs_in(first, end);
  std::uint64_t total = 0;
  for (const StyleRun* run = lo; run != hi; ++run) {
    const std::uint64_t from = std::max<std::uint64_t>(run->first, first);
    const std::uint64_t to = std::min<std::uint64_t>(static_cast<std::uint64_t>(run->last) + 1U, end);
    total += to - from;
  }
  return static_cast<std::uint32_t>(total);
}

std::optional<StyleSlot> RowCells::find_style(std::uint32_t col) const noexcept {
  if (const Cell* cell = find(col); cell != nullptr) {
    return StyleSlot{cell->xf_index, cell->has_explicit_xf};
  }
  if (runs_ != nullptr) {
    const std::size_t index = run_lower(col);
    if (index < runs_->size() && (*runs_)[index].first <= col) {
      return StyleSlot{(*runs_)[index].xf, true};
    }
  }
  return std::nullopt;
}

std::size_t RowCells::set_runs(std::uint32_t first, std::uint64_t end, std::uint32_t xf) {
  const std::uint64_t stop = GridStop(end);
  if (first >= stop) {
    return 0U;
  }
  const std::size_t before = run_total();
  std::vector<StyleRun> pieces;
  std::uint64_t at = first;
  for (std::size_t i = lower_index(first); i < cols_.size() && cols_[i] < stop; ++i) {
    if (cols_[i] > at) {
      pieces.push_back(StyleRun{static_cast<std::uint16_t>(at), static_cast<std::uint16_t>(cols_[i] - 1U), xf});
    }
    at = static_cast<std::uint64_t>(cols_[i]) + 1U;
  }
  if (at < stop) {
    pieces.push_back(StyleRun{static_cast<std::uint16_t>(at), static_cast<std::uint16_t>(stop - 1U), xf});
  }
  if (runs_ == nullptr) {
    if (pieces.empty()) {
      return 0U;
    }
    runs_ = std::make_unique<std::vector<StyleRun>>();
  }
  const std::size_t lo = split_runs_at(first);
  const std::size_t hi = split_runs_at(stop);
  splice_runs(*runs_, lo, hi, pieces.data(), pieces.data() + pieces.size());
  normalize_runs();
  assert(invariants_hold());
  const std::size_t after = run_total();
  return after > before ? after - before : 0U;
}

void RowCells::clear_runs(std::uint32_t first, std::uint64_t end) {
  const std::uint64_t stop = GridStop(end);
  if (runs_ == nullptr || first >= stop) {
    return;
  }
  const std::size_t lo = split_runs_at(first);
  const std::size_t hi = split_runs_at(stop);
  runs_->erase(runs_->begin() + static_cast<std::ptrdiff_t>(lo), runs_->begin() + static_cast<std::ptrdiff_t>(hi));
  normalize_runs();
  assert(invariants_hold());
}

std::size_t RowCells::remap_runs(std::uint32_t first, std::uint64_t end, const XfMap& map) {
  const std::uint64_t stop = GridStop(end);
  if (runs_ == nullptr || first >= stop) {
    return 0U;
  }
  const std::size_t before = run_total();
  const std::size_t lo = split_runs_at(first);
  const std::size_t hi = split_runs_at(stop);
  for (std::size_t i = lo; i < hi; ++i) {
    (*runs_)[i].xf = map((*runs_)[i].xf);
  }
  normalize_runs();
  assert(invariants_hold());
  const std::size_t after = run_total();
  return after > before ? after - before : 0U;
}

std::optional<std::uint32_t> RowCells::demote(std::uint32_t col) {
  const std::size_t index = index_of(col);
  if (index == kNpos || HasContent(cells_[index])) {
    return std::nullopt;
  }
  const Cell& cell = cells_[index];
  std::optional<std::uint32_t> owned;
  if (cell.xf_index != 0U || cell.has_explicit_xf) {
    owned = cell.xf_index;
  }
  cols_.erase(cols_.begin() + static_cast<std::ptrdiff_t>(index));
  cells_.erase(cells_.begin() + static_cast<std::ptrdiff_t>(index));
  return owned;
}

void RowCells::shift_cols_from(std::uint32_t col, std::uint32_t count) {
  const std::size_t first = lower_index(col);
  std::size_t kept = cols_.size();
  for (std::size_t i = first; i < cols_.size(); ++i) {
    const std::uint64_t shifted = static_cast<std::uint64_t>(cols_[i]) + count;
    if (shifted >= a1::kMaxCols) {
      kept = i;
      break;
    }
    cols_[i] = static_cast<std::uint16_t>(shifted);
  }
  cols_.resize(kept);
  cells_.erase(cells_.begin() + static_cast<std::ptrdiff_t>(kept), cells_.end());
  if (runs_ != nullptr && count != 0U) {
    std::vector<StyleRun>& list = *runs_;
    for (std::size_t i = split_runs_at(col); i < list.size(); ++i) {
      const std::uint64_t shifted_first = static_cast<std::uint64_t>(list[i].first) + count;
      if (shifted_first >= a1::kMaxCols) {
        list.resize(i);
        break;
      }
      const std::uint64_t shifted_last =
          std::min<std::uint64_t>(static_cast<std::uint64_t>(list[i].last) + count, a1::kMaxCols - 1U);
      list[i].first = static_cast<std::uint16_t>(shifted_first);
      list[i].last = static_cast<std::uint16_t>(shifted_last);
    }
    normalize_runs();
  }
  assert(invariants_hold());
}

void RowCells::erase(std::uint32_t col) {
  const std::size_t index = index_of(col);
  if (index != kNpos) {
    cols_.erase(cols_.begin() + static_cast<std::ptrdiff_t>(index));
    cells_.erase(cells_.begin() + static_cast<std::ptrdiff_t>(index));
  }
  clear_runs(col, static_cast<std::uint64_t>(col) + 1U);
  assert(invariants_hold());
}

void RowCells::erase_cols(std::uint32_t col, std::uint32_t count) {
  const std::uint64_t band_end = static_cast<std::uint64_t>(col) + count;
  const std::size_t first = lower_index(col);
  const std::size_t last = std::max(first, lower_index(band_end));
  cols_.erase(cols_.begin() + static_cast<std::ptrdiff_t>(first), cols_.begin() + static_cast<std::ptrdiff_t>(last));
  cells_.erase(cells_.begin() + static_cast<std::ptrdiff_t>(first), cells_.begin() + static_cast<std::ptrdiff_t>(last));
  for (std::size_t i = first; i < cols_.size(); ++i) {
    cols_[i] = static_cast<std::uint16_t>(cols_[i] - count);
  }
  if (runs_ != nullptr && count != 0U) {
    std::vector<StyleRun>& list = *runs_;
    const std::size_t lo = split_runs_at(col);
    const std::size_t hi = split_runs_at(band_end);
    list.erase(list.begin() + static_cast<std::ptrdiff_t>(lo), list.begin() + static_cast<std::ptrdiff_t>(hi));
    for (std::size_t i = lo; i < list.size(); ++i) {
      list[i].first = static_cast<std::uint16_t>(list[i].first - count);
      list[i].last = static_cast<std::uint16_t>(list[i].last - count);
    }
    normalize_runs();
  }
  assert(invariants_hold());
}

void RowCells::erase_in(std::uint32_t first, std::uint64_t end) {
  const std::size_t lo = lower_index(first);
  const std::size_t hi = std::max(lo, lower_index(end));
  cols_.erase(cols_.begin() + static_cast<std::ptrdiff_t>(lo), cols_.begin() + static_cast<std::ptrdiff_t>(hi));
  cells_.erase(cells_.begin() + static_cast<std::ptrdiff_t>(lo), cells_.begin() + static_cast<std::ptrdiff_t>(hi));
  clear_runs(first, end);
  assert(invariants_hold());
}

void RowCells::move_in_to(std::uint32_t first, std::uint64_t end, RowCells& dest) {
  dest.erase_in(first, end);
  const std::size_t lo = lower_index(first);
  const std::size_t hi = std::max(lo, lower_index(end));
  for (std::size_t i = lo; i < hi; ++i) {
    dest.ensure(cols_[i]) = std::move(cells_[i]);
  }
  const std::uint64_t stop = GridStop(end);
  if (runs_ != nullptr && first < stop) {
    const std::size_t run_lo = split_runs_at(first);
    const std::size_t run_hi = split_runs_at(stop);
    if (run_hi > run_lo) {
      if (dest.runs_ == nullptr) {
        dest.runs_ = std::make_unique<std::vector<StyleRun>>();
      }
      // `dest` holds no run inside the band any more, so this is the slot.
      const std::size_t at = dest.split_runs_at(first);
      splice_runs(*dest.runs_, at, at, runs_->data() + run_lo, runs_->data() + run_hi);
      runs_->erase(runs_->begin() + static_cast<std::ptrdiff_t>(run_lo),
                   runs_->begin() + static_cast<std::ptrdiff_t>(run_hi));
      dest.normalize_runs();
    }
    normalize_runs();
  }
  erase_in(first, end);
  assert(dest.invariants_hold());
}

std::size_t RowCells::run_lower(std::uint64_t col) const noexcept {
  return static_cast<std::size_t>(std::lower_bound(runs_->begin(), runs_->end(), col,
                                                   [](const StyleRun& run, std::uint64_t value) {
                                                     return static_cast<std::uint64_t>(run.last) < value;
                                                   }) -
                                  runs_->begin());
}

std::size_t RowCells::split_runs_at(std::uint64_t col) {
  std::vector<StyleRun>& list = *runs_;
  const std::size_t index = run_lower(col);
  if (index == list.size() || list[index].first >= col) {
    return index;
  }
  const StyleRun right{static_cast<std::uint16_t>(col), list[index].last, list[index].xf};
  list[index].last = static_cast<std::uint16_t>(col - 1U);
  list.insert(list.begin() + static_cast<std::ptrdiff_t>(index) + 1, right);
  return index + 1U;
}

void RowCells::normalize_runs() {
  if (runs_ == nullptr) {
    return;
  }
  std::vector<StyleRun>& list = *runs_;
  std::size_t out = 0;
  for (std::size_t i = 0; i < list.size(); ++i) {
    if (out != 0U && list[out - 1U].xf == list[i].xf && list[out - 1U].last + 1U == list[i].first) {
      list[out - 1U].last = list[i].last;
    } else {
      list[out++] = list[i];
    }
  }
  list.resize(out);
  if (list.empty()) {
    runs_.reset();
  }
}

std::optional<std::uint32_t> RowCells::take_run_column(std::uint32_t col) {
  if (runs_ == nullptr) {
    return std::nullopt;
  }
  const std::size_t index = run_lower(col);
  if (index == runs_->size() || (*runs_)[index].first > col) {
    return std::nullopt;
  }
  const std::uint32_t xf = (*runs_)[index].xf;
  clear_runs(col, static_cast<std::uint64_t>(col) + 1U);
  return xf;
}

bool RowCells::invariants_hold() const noexcept {
  if (runs_ == nullptr) {
    return true;
  }
  const std::vector<StyleRun>& list = *runs_;
  if (list.empty()) {
    return false;
  }
  for (std::size_t i = 0; i < list.size(); ++i) {
    if (list[i].first > list[i].last || list[i].last >= a1::kMaxCols) {
      return false;
    }
    if (i != 0U) {
      const StyleRun& prev = list[i - 1U];
      if (prev.last >= list[i].first || (prev.last + 1U == list[i].first && prev.xf == list[i].xf)) {
        return false;
      }
    }
  }
  std::size_t j = 0;
  for (const std::uint16_t col : cols_) {
    while (j < list.size() && list[j].last < col) {
      ++j;
    }
    if (j < list.size() && list[j].first <= col) {
      return false;
    }
  }
  return true;
}

}  // namespace formulon
