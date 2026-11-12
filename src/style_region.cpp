// Range style edits over row, column and per-row run storage: remap, clear,
// snapshot and restore.
//
// @size-budget: 28 KB

#include "style_region.h"

#include <algorithm>
#include <limits>
#include <string>
#include <utility>

#include "cell.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "style_resolve.h"
#include "utils/index_sort.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"

namespace formulon {
namespace {

constexpr std::uint8_t kSnapshotMagic[4] = {'F', 'M', 'R', 'S'};
constexpr std::uint8_t kSnapshotVersionCells = 1U;
constexpr std::uint8_t kSnapshotVersion = 2U;
constexpr std::uint8_t kEntryCell = 0U;
constexpr std::uint8_t kEntryRun = 1U;
constexpr std::uint8_t kCapturesColumns = 1U;
constexpr std::uint8_t kCapturesRows = 2U;

// A run of columns sharing one column-style state.
struct ColumnSegment {
  std::uint32_t first = 0;
  std::uint32_t last = 0;
  bool styled = false;
  std::uint32_t xf = 0;
};

// A styled row and its xf.
struct RowStyle {
  std::uint32_t row = 0;
  std::uint32_t xf = 0;
};

struct CapturedCell {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::uint32_t xf = 0;
  bool explicit_xf = false;
};

// A captured style run, inclusive columns.
struct CapturedRun {
  std::uint32_t row = 0;
  std::uint32_t first = 0;
  std::uint32_t last = 0;
  std::uint32_t xf = 0;
};

bool cell_less(const StyleCell& a, const StyleCell& b) {
  return a.row != b.row ? a.row < b.row : a.col < b.col;
}

bool covers_all_rows(const StyleRect& rect) {
  return rect.first_row == 0U && rect.last_row == Sheet::kMaxRows - 1U;
}

bool covers_all_cols(const StyleRect& rect) {
  return rect.first_col == 0U && rect.last_col == Sheet::kMaxCols - 1U;
}

// Column-style segments tiling `[first, last]`, gaps included as unstyled.
std::vector<ColumnSegment> column_segments(const SheetLayout& layout, std::uint32_t first, std::uint32_t last) {
  std::vector<ColumnSegment> styled;
  for (const ColumnLayout& entry : layout.columns) {
    if (entry.has_style && entry.last >= first && entry.first <= last) {
      styled.push_back({std::max(entry.first, first), std::min(entry.last, last), true, entry.style_xf});
    }
  }
  sort_by_index(styled, [](const ColumnSegment& lhs, const ColumnSegment& rhs) { return lhs.first < rhs.first; });
  std::vector<ColumnSegment> out;
  std::uint64_t cursor = first;
  for (const ColumnSegment& seg : styled) {
    if (cursor < seg.first) {
      out.push_back({static_cast<std::uint32_t>(cursor), seg.first - 1U, false, 0U});
    }
    out.push_back(seg);
    cursor = static_cast<std::uint64_t>(seg.last) + 1U;
  }
  if (cursor <= last) {
    out.push_back({static_cast<std::uint32_t>(cursor), last, false, 0U});
  }
  return out;
}

// Styled rows of `[first, last]`, ascending (the layout keeps rows ordered).
std::vector<RowStyle> styled_rows(const SheetLayout& layout, std::uint32_t first, std::uint32_t last) {
  std::vector<RowStyle> out;
  for (auto it = row_override_lower_bound(layout, first); it != layout.row_overrides.end() && it->row <= last; ++it) {
    if (it->has_style) {
      out.push_back({it->row, it->style_xf});
    }
  }
  return out;
}

// Every stored cell of `rect` with its style state; `explicit_xf` here means
// the cell owns its style.
std::vector<CapturedCell> stored_cells(const Sheet& sheet, const StyleRect& rect) {
  std::vector<CapturedCell> out;
  for (const auto& [row, cells] : sheet.rows()) {
    if (row < rect.first_row || row > rect.last_row || !cells.has_any()) {
      continue;
    }
    for (const auto& [col, cell] : cells.entries_in(rect.first_col, static_cast<std::uint64_t>(rect.last_col) + 1U)) {
      out.push_back({row, col, cell.xf_index, cell_owns_style(cell)});
    }
  }
  return out;
}

// Rows of `rect` holding a stored cell or a style run inside its columns, ascending.
std::vector<std::uint32_t> rows_holding_styles(const Sheet& sheet, const StyleRect& rect) {
  std::vector<std::uint32_t> out;
  for (const auto& [row, cells] : sheet.rows()) {
    if (row >= rect.first_row && row <= rect.last_row &&
        cells.any_in(rect.first_col, static_cast<std::uint64_t>(rect.last_col) + 1U)) {
      out.push_back(row);
    }
  }
  sort_ascending(out);
  return out;
}

// The runs of `cells` inside `rect`'s columns, clipped to them.
template <typename Fn>
void for_each_clipped_run(const RowCells& cells, const StyleRect& rect, Fn&& fn) {
  const auto [lo, hi] = cells.runs_in(rect.first_col, static_cast<std::uint64_t>(rect.last_col) + 1U);
  for (const StyleRun* run = lo; run != hi; ++run) {
    fn(std::max<std::uint32_t>(run->first, rect.first_col), std::min<std::uint32_t>(run->last, rect.last_col), run->xf);
  }
}

// Decides the style of one row of a rewrite from its current style; a result
// with `styled == false` leaves the row unstyled.
struct RowDecision {
  bool styled = false;
  std::uint32_t xf = 0;
};
using RowDecide = RowDecision (*)(void* ctx, std::uint32_t row, RowDecision current);

// Applies the decided style to one row and keeps the entry unless it became
// an erasable override.
void finish_row(RowLayout& entry, RowDecide decide, void* ctx, bool default_hidden, bool existed_before,
                std::vector<RowLayout>& out) {
  const RowDecision decision = decide(ctx, entry.row, {entry.has_style, entry.style_xf});
  entry.has_style = decision.styled;
  entry.style_xf = decision.styled ? decision.xf : 0U;
  if (!should_erase_row_override(entry, default_hidden, existed_before)) {
    out.push_back(entry);
  }
}

// Rewrites every row of `[first, last]`; use for operations that really do
// materialize a style on every row. New entries inherit the sheet's default
// visibility, and pre-existing visible exceptions survive a style clear.
void rewrite_row_styles(SheetLayout& layout, std::uint32_t first, std::uint32_t last, bool default_hidden,
                        RowDecide decide, void* ctx) {
  std::vector<RowLayout>& rows = layout.row_overrides;
  std::vector<RowLayout> out;
  out.reserve(rows.size());
  std::size_t next = static_cast<std::size_t>(row_override_lower_bound(layout, first) - rows.begin());
  out.insert(out.end(), rows.begin(), rows.begin() + static_cast<std::ptrdiff_t>(next));
  for (std::uint64_t row = first; row <= last; ++row) {
    RowLayout entry;
    entry.row = static_cast<std::uint32_t>(row);
    bool existed_before = false;
    if (next < rows.size() && rows[next].row == row) {
      entry = rows[next++];
      existed_before = true;
    } else {
      entry.hidden = default_hidden;
    }
    finish_row(entry, decide, ctx, default_hidden, existed_before, out);
  }
  out.insert(out.end(), rows.begin() + static_cast<std::ptrdiff_t>(next), rows.end());
  rows = std::move(out);
}

// Rewrites only existing row overrides and any explicit rows supplied by a
// capture. The two ordered inputs are merged into one pass; absent rows are
// materialized only when the callback assigns them state.
void rewrite_sparse_row_styles(SheetLayout& layout, std::uint32_t first, std::uint32_t last, bool default_hidden,
                               RowDecide decide, void* ctx, const std::vector<RowStyle>* additional_rows = nullptr) {
  std::vector<RowLayout>& rows = layout.row_overrides;
  std::vector<RowLayout> out;
  out.reserve(rows.size() + (additional_rows == nullptr ? 0U : additional_rows->size()));
  std::size_t next = static_cast<std::size_t>(row_override_lower_bound(layout, first) - rows.begin());
  out.insert(out.end(), rows.begin(), rows.begin() + static_cast<std::ptrdiff_t>(next));
  std::size_t extra = 0U;
  if (additional_rows != nullptr) {
    extra = static_cast<std::size_t>(
        std::lower_bound(additional_rows->begin(), additional_rows->end(), first,
                         [](const RowStyle& entry, std::uint32_t key) { return entry.row < key; }) -
        additional_rows->begin());
  }

  while (true) {
    const bool has_existing = next < rows.size() && rows[next].row <= last;
    const bool has_additional =
        additional_rows != nullptr && extra < additional_rows->size() && (*additional_rows)[extra].row <= last;
    if (!has_existing && !has_additional) {
      break;
    }
    const std::uint32_t row = !has_additional || (has_existing && rows[next].row < (*additional_rows)[extra].row)
                                  ? rows[next].row
                                  : (*additional_rows)[extra].row;
    RowLayout entry;
    entry.row = row;
    bool existed_before = false;
    if (has_existing && rows[next].row == row) {
      entry = rows[next++];
      existed_before = true;
    } else {
      entry.hidden = default_hidden;
    }
    if (has_additional && (*additional_rows)[extra].row == row) {
      ++extra;
    }
    finish_row(entry, decide, ctx, default_hidden, existed_before, out);
  }
  out.insert(out.end(), rows.begin() + static_cast<std::ptrdiff_t>(next), rows.end());
  rows = std::move(out);
}

void append_u32(std::vector<std::uint8_t>& out, std::uint32_t value) {
  for (int shift = 0; shift < 32; shift += 8) {
    out.push_back(static_cast<std::uint8_t>((value >> shift) & 0xFFU));
  }
}

void put_u32_at(std::vector<std::uint8_t>& out, std::size_t pos, std::uint32_t value) {
  for (int shift = 0; shift < 32; shift += 8) {
    out[pos++] = static_cast<std::uint8_t>((value >> shift) & 0xFFU);
  }
}

class ByteReader {
 public:
  ByteReader(const std::uint8_t* bytes, std::size_t len) : bytes_(bytes), len_(len) {}
  bool u8(std::uint8_t& out) {
    if (pos_ + 1U > len_) {
      return false;
    }
    out = bytes_[pos_++];
    return true;
  }
  bool u32(std::uint32_t& out) {
    if (pos_ + 4U > len_) {
      return false;
    }
    out = 0U;
    for (int shift = 0; shift < 32; shift += 8) {
      out |= static_cast<std::uint32_t>(bytes_[pos_++]) << shift;
    }
    return true;
  }
  bool at_end() const { return pos_ == len_; }

 private:
  const std::uint8_t* bytes_;
  std::size_t len_;
  std::size_t pos_ = 0;
};

struct RemapRows {
  const XfMap* map;
  bool style_bare;
  std::uint32_t mapped_default;
};

struct RestoreRows {
  const std::vector<RowStyle>* rows;
  std::size_t next;
};

}  // namespace

void set_column_style_span(SheetLayout& layout, std::uint32_t first, std::uint32_t last, bool styled,
                           std::uint32_t xf_index) {
  overlay_column_span(layout, first, last, [styled, xf_index](ColumnLayout& entry) {
    entry.has_style = styled;
    entry.style_xf = styled ? xf_index : 0U;
  });
  if (!styled) {
    drop_stateless_columns(layout);
  }
}

void set_row_style_span(SheetLayout& layout, std::uint32_t first, std::uint32_t last, std::uint32_t xf_index,
                        bool default_hidden) {
  rewrite_row_styles(
      layout, first, last, default_hidden,
      [](void* ctx, std::uint32_t, RowDecision) { return RowDecision{true, *static_cast<std::uint32_t*>(ctx)}; },
      &xf_index);
}

void clear_row_style_span(SheetLayout& layout, std::uint32_t first, std::uint32_t last, bool default_hidden) {
  rewrite_sparse_row_styles(
      layout, first, last, default_hidden, [](void*, std::uint32_t, RowDecision) { return RowDecision{}; }, nullptr);
}

Expected<void, Error> check_styled_range_area(const StyleRect& rect) {
  const std::uint64_t area = (static_cast<std::uint64_t>(rect.last_row - rect.first_row) + 1U) *
                             (static_cast<std::uint64_t>(rect.last_col - rect.first_col) + 1U);
  if (!covers_all_rows(rect) && !covers_all_cols(rect) && area > kMaxStyledRangeCells) {
    return make_error(FormulonErrorCode::kSecurityResourceLimit, "range is past the styled-range cell ceiling",
                      limit_context("styled_range_cells", area, kMaxStyledRangeCells));
  }
  return Expected<void, Error>::Ok();
}

std::size_t count_cell_xf_indices(const Sheet& sheet, const StyleRect& rect) {
  std::size_t total = 0U;
  for (const auto& [row, cells] : sheet.rows()) {
    if (row < rect.first_row || row > rect.last_row || !cells.has_any()) {
      continue;
    }
    for (const auto& [col, cell] : cells.entries_in(rect.first_col, static_cast<std::uint64_t>(rect.last_col) + 1U)) {
      (void)col;
      total += cell.xf_index != 0U ? 1U : 0U;
    }
    for_each_clipped_run(cells, rect, [&total](std::uint32_t first, std::uint32_t last, std::uint32_t xf) {
      total += xf != 0U ? static_cast<std::size_t>(last - first) + 1U : 0U;
    });
  }
  return total;
}

std::size_t copy_cell_xf_indices(const Sheet& sheet, const StyleRect& rect, StyleCellXf* out, std::size_t capacity) {
  std::size_t written = 0U;
  for (const std::uint32_t row : rows_holding_styles(sheet, rect)) {
    if (written == capacity) {
      break;
    }
    const RowCells& cells = sheet.rows().find(row)->second;
    const RowCells::EntryRange entries =
        cells.entries_in(rect.first_col, static_cast<std::uint64_t>(rect.last_col) + 1U);
    const auto [run_lo, run_hi] = cells.runs_in(rect.first_col, static_cast<std::uint64_t>(rect.last_col) + 1U);
    auto entry = entries.begin();
    const StyleRun* run = run_lo;
    // Stored cells and runs are disjoint, so one merge walk gives column order.
    while (written < capacity && (entry != entries.end() || run != run_hi)) {
      if (run == run_hi || (entry != entries.end() && (*entry).col < run->first)) {
        if ((*entry).cell.xf_index != 0U) {
          out[written++] = {row, (*entry).col, (*entry).cell.xf_index};
        }
        ++entry;
        continue;
      }
      if (run->xf != 0U) {
        const std::uint32_t first = std::max<std::uint32_t>(run->first, rect.first_col);
        const std::uint32_t last = std::min<std::uint32_t>(run->last, rect.last_col);
        for (std::uint64_t col = first; col <= last && written < capacity; ++col) {
          out[written++] = {row, static_cast<std::uint32_t>(col), run->xf};
        }
      }
      ++run;
    }
  }
  return written;
}

std::vector<StyleCell> show_phonetic_cells(const Sheet& sheet, const StyleRect& rect) {
  std::vector<StyleCell> out;
  for (const auto& [row, cells] : sheet.rows()) {
    if (row < rect.first_row || row > rect.last_row || !cells.has_any()) {
      continue;
    }
    for (const auto& [col, cell] : cells.entries_in(rect.first_col, static_cast<std::uint64_t>(rect.last_col) + 1U)) {
      if (cell.show_phonetic) {
        out.push_back({row, col});
      }
    }
  }
  sort_by_index(out,
                [](const StyleCell& a, const StyleCell& b) { return a.row != b.row ? a.row < b.row : a.col < b.col; });
  return out;
}

std::vector<std::uint32_t> range_style_bases(const Sheet& sheet, const StyleRect& rect,
                                             std::vector<StyleCell> excluded) {
  const auto cell_equal = [](const StyleCell& a, const StyleCell& b) { return a.row == b.row && a.col == b.col; };
  excluded.erase(std::remove_if(excluded.begin(), excluded.end(),
                                [&rect](const StyleCell& c) {
                                  return c.row < rect.first_row || c.row > rect.last_row || c.col < rect.first_col ||
                                         c.col > rect.last_col;
                                }),
                 excluded.end());
  sort_by_index(excluded, cell_less);
  excluded.erase(std::unique(excluded.begin(), excluded.end(), cell_equal), excluded.end());

  const std::vector<RowStyle> rows = styled_rows(sheet.layout(), rect.first_row, rect.last_row);
  const std::vector<ColumnSegment> cols = column_segments(sheet.layout(), rect.first_col, rect.last_col);
  // Coordinates that hide the axis style beneath them, per styled row and per
  // column segment across the unstyled rows: owning cells, run columns and
  // excluded cells.
  std::vector<std::uint64_t> owned_in_row(rows.size(), 0U);
  std::vector<std::uint64_t> owned_in_col(cols.size(), 0U);
  const auto count_covered = [&](std::uint32_t row, std::uint32_t first, std::uint32_t last) {
    const auto row_it = std::lower_bound(rows.begin(), rows.end(), row,
                                         [](const RowStyle& entry, std::uint32_t key) { return entry.row < key; });
    if (row_it != rows.end() && row_it->row == row) {
      owned_in_row[static_cast<std::size_t>(row_it - rows.begin())] += static_cast<std::uint64_t>(last - first) + 1U;
      return;
    }
    auto col_it = std::upper_bound(cols.begin(), cols.end(), first,
                                   [](std::uint32_t key, const ColumnSegment& seg) { return key < seg.first; });
    for (--col_it; col_it != cols.end() && col_it->first <= last; ++col_it) {
      owned_in_col[static_cast<std::size_t>(col_it - cols.begin())] +=
          static_cast<std::uint64_t>(std::min(col_it->last, last) - std::max(col_it->first, first)) + 1U;
    }
  };
  // Excluded cells in `[first, last]` of `row`.
  const auto excluded_in = [&excluded](std::uint32_t row, std::uint32_t first, std::uint32_t last) {
    const auto lo = std::lower_bound(excluded.begin(), excluded.end(), StyleCell{row, first}, cell_less);
    const auto hi = std::upper_bound(lo, excluded.end(), StyleCell{row, last}, cell_less);
    return static_cast<std::uint64_t>(hi - lo);
  };
  std::vector<std::uint32_t> out;
  std::vector<StyleCell> owners;
  for (const CapturedCell& cell : stored_cells(sheet, rect)) {
    if (!cell.explicit_xf) {
      continue;
    }
    const StyleCell at{cell.row, cell.col};
    if (excluded.empty()) {
      out.push_back(cell.xf);
    } else {
      if (!std::binary_search(excluded.begin(), excluded.end(), at, cell_less)) {
        out.push_back(cell.xf);
      }
      owners.push_back(at);
    }
    count_covered(at.row, at.col, at.col);
  }
  for (const auto& entry : sheet.rows()) {
    const std::uint32_t row = entry.first;
    if (row < rect.first_row || row > rect.last_row) {
      continue;
    }
    for_each_clipped_run(entry.second, rect, [&](std::uint32_t first, std::uint32_t last, std::uint32_t xf) {
      if (excluded_in(row, first, last) < static_cast<std::uint64_t>(last - first) + 1U) {
        out.push_back(xf);
      }
      count_covered(row, first, last);
    });
  }
  if (!excluded.empty()) {
    sort_by_index(owners, cell_less);
    for (const StyleCell& cell : excluded) {
      if (std::binary_search(owners.begin(), owners.end(), cell, cell_less)) {
        continue;
      }
      // A run column already counted as an owner.
      const auto row_it = sheet.rows().find(cell.row);
      if (row_it != sheet.rows().end() && row_it->second.find(cell.col) == nullptr &&
          row_it->second.find_style(cell.col).has_value()) {
        continue;
      }
      count_covered(cell.row, cell.col, cell.col);
    }
  }
  const std::uint64_t width = static_cast<std::uint64_t>(rect.last_col - rect.first_col) + 1U;
  for (std::size_t i = 0; i < rows.size(); ++i) {
    if (owned_in_row[i] < width) {
      out.push_back(rows[i].xf);
    }
  }
  const std::uint64_t bare_rows = static_cast<std::uint64_t>(rect.last_row - rect.first_row) + 1U - rows.size();
  for (std::size_t i = 0; i < cols.size(); ++i) {
    const std::uint64_t area = bare_rows * (static_cast<std::uint64_t>(cols[i].last - cols[i].first) + 1U);
    if (owned_in_col[i] < area) {
      out.push_back(cols[i].styled ? cols[i].xf : 0U);
    }
  }
  sort_ascending(out);
  out.erase(std::unique(out.begin(), out.end()), out.end());
  return out;
}

Expected<void, Error> remap_range_style(Sheet& sheet, const StyleRect& rect, const XfMap& map) {
  RETURN_IF_ERROR(check_styled_range_area(rect));
  const bool all_rows = covers_all_rows(rect);
  const bool all_cols = covers_all_cols(rect);
  const std::uint64_t col_end = static_cast<std::uint64_t>(rect.last_col) + 1U;
  std::vector<CapturedCell> owned = stored_cells(sheet, rect);
  owned.erase(std::remove_if(owned.begin(), owned.end(), [](const CapturedCell& cell) { return !cell.explicit_xf; }),
              owned.end());
  sort_by_index(owned, [](const CapturedCell& a, const CapturedCell& b) {
    return a.row != b.row ? a.row < b.row : a.col < b.col;
  });
  std::vector<std::uint32_t> run_rows;
  for (const auto& [row, cells] : sheet.rows()) {
    if (row >= rect.first_row && row <= rect.last_row && cells.run_count_in(rect.first_col, col_end) != 0U) {
      run_rows.push_back(row);
    }
  }
  sort_ascending(run_rows);

  const bool default_hidden = sheet.format_defaults().zero_height;
  const std::uint32_t mapped_default = map(0U);
  const SheetLayout& layout = sheet.layout();
  const std::vector<ColumnSegment> pre_cols = column_segments(layout, rect.first_col, rect.last_col);
  const std::vector<RowStyle> pre_rows = styled_rows(layout, rect.first_row, rect.last_row);

  std::vector<ColumnSegment> post_cols = pre_cols;
  if (all_rows) {
    for (ColumnSegment& seg : post_cols) {
      seg.xf = map(seg.styled ? seg.xf : 0U);
      seg.styled = seg.styled || mapped_default != 0U;
    }
  }
  const bool style_bare_rows = all_cols && !all_rows && mapped_default != 0U;
  // A row band leaves column styles in place, so bare rows under a remapped column need cell writes.
  const bool bare_rows_cross_remapped_cols =
      all_cols && !all_rows && std::any_of(pre_cols.begin(), pre_cols.end(), [&map](const ColumnSegment& seg) {
        return seg.styled && map(seg.xf) != seg.xf;
      });

  // Cell writes are runs `(row, first, last, xf)` of column segments. A plain
  // range writes every segment; an axis selection only those where the
  // surviving row-over-column precedence would resolve to the wrong xf.
  struct CellRun {
    std::uint32_t row;
    std::uint32_t first;
    std::uint32_t last;
    std::uint32_t xf;
  };
  std::vector<CellRun> runs;
  const auto plan_row = [&](std::uint32_t row, const RowStyle* pre_row) {
    std::uint32_t post_row_xf = 0U;
    bool post_row_styled = pre_row != nullptr;
    if (pre_row != nullptr) {
      post_row_xf = all_cols ? map(pre_row->xf) : pre_row->xf;
    } else if (style_bare_rows) {
      post_row_styled = true;
      post_row_xf = mapped_default;
    }
    for (std::size_t i = 0; i < pre_cols.size(); ++i) {
      const ColumnSegment& pre = pre_cols[i];
      const ColumnSegment& post = post_cols[i];
      const std::uint32_t target = map(pre_row != nullptr ? pre_row->xf : (pre.styled ? pre.xf : 0U));
      const std::uint32_t after = post_row_styled ? post_row_xf : (post.styled ? post.xf : 0U);
      if ((all_rows || all_cols) && after == target) {
        continue;
      }
      runs.push_back({row, pre.first, pre.last, target});
    }
  };
  if ((all_rows || all_cols) && !style_bare_rows && !bare_rows_cross_remapped_cols) {
    // Only styled row overrides can require cell writes.
    for (const RowStyle& pre_row : pre_rows) {
      plan_row(pre_row.row, &pre_row);
    }
  } else {
    std::size_t next = 0;
    for (std::uint64_t row = rect.first_row; row <= rect.last_row; ++row) {
      const RowStyle* pre_row = nullptr;
      if (next < pre_rows.size() && pre_rows[next].row == row) {
        pre_row = &pre_rows[next++];
      }
      plan_row(static_cast<std::uint32_t>(row), pre_row);
    }
  }

  SheetLayout& mutable_layout = sheet.mutable_layout();
  if (all_rows) {
    for (const ColumnSegment& seg : post_cols) {
      if (seg.styled) {
        set_column_style_span(mutable_layout, seg.first, seg.last, true, seg.xf);
      }
    }
  }
  if (all_cols) {
    RemapRows ctx{&map, style_bare_rows, mapped_default};
    const RowDecide remap_row = [](void* raw, std::uint32_t, RowDecision current) {
      const auto* remap = static_cast<RemapRows*>(raw);
      if (current.styled) {
        return RowDecision{true, (*remap->map)(current.xf)};
      }
      return RowDecision{remap->style_bare, remap->mapped_default};
    };
    if (style_bare_rows) {
      rewrite_row_styles(mutable_layout, rect.first_row, rect.last_row, default_hidden, remap_row, &ctx);
    } else {
      rewrite_sparse_row_styles(mutable_layout, rect.first_row, rect.last_row, default_hidden, remap_row, &ctx);
    }
  }

  // Per row, against the rewritten layout: the runs already there (mapped
  // once), then the planned pieces around them, then the stored owners.
  std::size_t run_row = 0U;
  std::size_t planned = 0U;
  std::size_t owner = 0U;
  while (run_row < run_rows.size() || planned < runs.size() || owner < owned.size()) {
    std::uint32_t row = std::numeric_limits<std::uint32_t>::max();
    if (run_row < run_rows.size()) {
      row = std::min(row, run_rows[run_row]);
    }
    if (planned < runs.size()) {
      row = std::min(row, runs[planned].row);
    }
    if (owner < owned.size()) {
      row = std::min(row, owned[owner].row);
    }
    if (run_row < run_rows.size() && run_rows[run_row] == row) {
      sheet.remap_style_runs(row, rect.first_col, col_end, map);
      ++run_row;
    }
    for (; planned < runs.size() && runs[planned].row == row; ++planned) {
      sheet.apply_planned_run(row, runs[planned].first, static_cast<std::uint64_t>(runs[planned].last) + 1U,
                              runs[planned].xf);
    }
    for (; owner < owned.size() && owned[owner].row == row; ++owner) {
      sheet.set_cell_xf_index(row, owned[owner].col, map(owned[owner].xf));
    }
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> set_range_xf_index(Sheet& sheet, const StyleRect& rect, std::uint32_t xf_index) {
  XfMap map;
  map.constant = xf_index;
  map.is_constant = true;
  return remap_range_style(sheet, rect, map);
}

std::vector<std::uint8_t> snapshot_range_style(const Sheet& sheet, const StyleRect& rect) {
  std::vector<std::uint8_t> out(std::begin(kSnapshotMagic), std::end(kSnapshotMagic));
  out.push_back(kSnapshotVersion);
  append_u32(out, rect.first_row);
  append_u32(out, rect.first_col);
  append_u32(out, rect.last_row);
  append_u32(out, rect.last_col);
  const bool all_rows = covers_all_rows(rect);
  const bool all_cols = covers_all_cols(rect);
  out.push_back(static_cast<std::uint8_t>((all_rows ? kCapturesColumns : 0U) | (all_cols ? kCapturesRows : 0U)));
  if (all_rows) {
    std::vector<ColumnSegment> styled;
    for (const ColumnSegment& seg : column_segments(sheet.layout(), rect.first_col, rect.last_col)) {
      if (seg.styled) {
        styled.push_back(seg);
      }
    }
    append_u32(out, static_cast<std::uint32_t>(styled.size()));
    for (const ColumnSegment& seg : styled) {
      append_u32(out, seg.first);
      append_u32(out, seg.last);
      append_u32(out, seg.xf);
    }
  }
  if (all_cols) {
    const std::vector<RowStyle> rows = styled_rows(sheet.layout(), rect.first_row, rect.last_row);
    append_u32(out, static_cast<std::uint32_t>(rows.size()));
    for (const RowStyle& row : rows) {
      append_u32(out, row.row);
      append_u32(out, row.xf);
    }
  }
  // Row groups: the count is patched once the rows with an entry are known.
  const std::size_t group_count_at = out.size();
  append_u32(out, 0U);
  std::uint32_t groups = 0U;
  for (const std::uint32_t row : rows_holding_styles(sheet, rect)) {
    const RowCells& cells = sheet.rows().find(row)->second;
    const std::size_t header_at = out.size();
    append_u32(out, row);
    append_u32(out, 0U);
    std::uint32_t entries = 0U;
    const RowCells::EntryRange stored =
        cells.entries_in(rect.first_col, static_cast<std::uint64_t>(rect.last_col) + 1U);
    const auto [run_lo, run_hi] = cells.runs_in(rect.first_col, static_cast<std::uint64_t>(rect.last_col) + 1U);
    auto entry = stored.begin();
    const StyleRun* run = run_lo;
    while (entry != stored.end() || run != run_hi) {
      if (run == run_hi || (entry != stored.end() && (*entry).col < run->first)) {
        const Cell& cell = (*entry).cell;
        if (cell.has_explicit_xf || cell.xf_index != 0U) {
          out.push_back(kEntryCell);
          append_u32(out, (*entry).col);
          append_u32(out, cell.xf_index);
          out.push_back(cell.has_explicit_xf ? 1U : 0U);
          ++entries;
        }
        ++entry;
        continue;
      }
      out.push_back(kEntryRun);
      append_u32(out, std::max<std::uint32_t>(run->first, rect.first_col));
      append_u32(out, std::min<std::uint32_t>(run->last, rect.last_col));
      append_u32(out, run->xf);
      ++entries;
      ++run;
    }
    if (entries == 0U) {
      out.resize(header_at);
      continue;
    }
    put_u32_at(out, header_at + 4U, entries);
    ++groups;
  }
  put_u32_at(out, group_count_at, groups);
  return out;
}

namespace {

// Decodes the per-coordinate cell list of a version-1 capture.
bool parse_cells_v1(ByteReader& in, const StyleRect& rect, std::size_t xf_count, std::vector<CapturedCell>& cells) {
  std::uint32_t count = 0;
  if (!in.u32(count)) {
    return false;
  }
  for (std::uint32_t i = 0; i < count; ++i) {
    CapturedCell cell;
    std::uint8_t explicit_xf = 0;
    if (!in.u32(cell.row) || !in.u32(cell.col) || !in.u32(cell.xf) || !in.u8(explicit_xf) || explicit_xf > 1U ||
        cell.row < rect.first_row || cell.row > rect.last_row || cell.col < rect.first_col ||
        cell.col > rect.last_col || cell.xf >= xf_count) {
      return false;
    }
    cell.explicit_xf = explicit_xf != 0U;
    cells.push_back(cell);
  }
  return true;
}

// Decodes the row groups of a version-2 capture: rows strictly ascending
// inside the rect, entries ascending and disjoint inside its columns.
bool parse_rows_v2(ByteReader& in, const StyleRect& rect, std::size_t xf_count, std::vector<CapturedCell>& cells,
                   std::vector<CapturedRun>& runs) {
  std::uint32_t groups = 0;
  if (!in.u32(groups)) {
    return false;
  }
  std::uint64_t min_row = rect.first_row;
  for (std::uint32_t g = 0; g < groups; ++g) {
    std::uint32_t row = 0;
    std::uint32_t entries = 0;
    if (!in.u32(row) || !in.u32(entries) || row < min_row || row > rect.last_row) {
      return false;
    }
    min_row = static_cast<std::uint64_t>(row) + 1U;
    std::uint64_t min_col = rect.first_col;
    for (std::uint32_t e = 0; e < entries; ++e) {
      std::uint8_t tag = 0;
      if (!in.u8(tag)) {
        return false;
      }
      if (tag == kEntryCell) {
        CapturedCell cell;
        cell.row = row;
        std::uint8_t explicit_xf = 0;
        if (!in.u32(cell.col) || !in.u32(cell.xf) || !in.u8(explicit_xf) || explicit_xf > 1U || cell.col < min_col ||
            cell.col > rect.last_col || cell.xf >= xf_count) {
          return false;
        }
        cell.explicit_xf = explicit_xf != 0U;
        min_col = static_cast<std::uint64_t>(cell.col) + 1U;
        cells.push_back(cell);
      } else if (tag == kEntryRun) {
        CapturedRun run;
        run.row = row;
        if (!in.u32(run.first) || !in.u32(run.last) || !in.u32(run.xf) || run.first > run.last || run.first < min_col ||
            run.last > rect.last_col || run.xf >= xf_count) {
          return false;
        }
        min_col = static_cast<std::uint64_t>(run.last) + 1U;
        runs.push_back(run);
      } else {
        return false;
      }
    }
  }
  return true;
}

// Decodes a `snapshot_range_style` capture of either version; false on any
// malformed field.
bool parse_capture(const std::uint8_t* bytes, std::size_t len, std::size_t xf_count, StyleRect& rect,
                   std::uint8_t& flags, std::vector<ColumnSegment>& cols, std::vector<RowStyle>& rows,
                   std::vector<CapturedCell>& cells, std::vector<CapturedRun>& runs) {
  if (bytes == nullptr || len < sizeof(kSnapshotMagic) ||
      !std::equal(std::begin(kSnapshotMagic), std::end(kSnapshotMagic), bytes)) {
    return false;
  }
  ByteReader in(bytes + sizeof(kSnapshotMagic), len - sizeof(kSnapshotMagic));
  std::uint8_t version = 0;
  if (!in.u8(version) || (version != kSnapshotVersionCells && version != kSnapshotVersion) || !in.u32(rect.first_row) ||
      !in.u32(rect.first_col) || !in.u32(rect.last_row) || !in.u32(rect.last_col) || !in.u8(flags) ||
      rect.first_row > rect.last_row || rect.first_col > rect.last_col || rect.last_row >= Sheet::kMaxRows ||
      rect.last_col >= Sheet::kMaxCols ||
      flags != ((covers_all_rows(rect) ? kCapturesColumns : 0U) | (covers_all_cols(rect) ? kCapturesRows : 0U))) {
    return false;
  }
  std::uint32_t count = 0;
  if ((flags & kCapturesColumns) != 0U) {
    if (!in.u32(count)) {
      return false;
    }
    for (std::uint32_t i = 0; i < count; ++i) {
      ColumnSegment seg;
      if (!in.u32(seg.first) || !in.u32(seg.last) || !in.u32(seg.xf) || seg.first > seg.last ||
          seg.first < rect.first_col || seg.last > rect.last_col || seg.xf >= xf_count) {
        return false;
      }
      cols.push_back(seg);
    }
  }
  if ((flags & kCapturesRows) != 0U) {
    if (!in.u32(count)) {
      return false;
    }
    for (std::uint32_t i = 0; i < count; ++i) {
      RowStyle row;
      if (!in.u32(row.row) || !in.u32(row.xf) || row.row < rect.first_row || row.row > rect.last_row ||
          (!rows.empty() && row.row <= rows.back().row) || row.xf >= xf_count) {
        return false;
      }
      rows.push_back(row);
    }
  }
  const bool entries_ok = version == kSnapshotVersionCells ? parse_cells_v1(in, rect, xf_count, cells)
                                                           : parse_rows_v2(in, rect, xf_count, cells, runs);
  return entries_ok && in.at_end();
}

}  // namespace

Expected<void, Error> restore_range_style(Sheet& sheet, const std::uint8_t* bytes, std::size_t len,
                                          std::size_t xf_count) {
  StyleRect rect;
  std::uint8_t flags = 0;
  std::vector<ColumnSegment> cols;
  std::vector<RowStyle> rows;
  std::vector<CapturedCell> cells;
  std::vector<CapturedRun> runs;
  if (!parse_capture(bytes, len, xf_count, rect, flags, cols, rows, cells, runs)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "bytes is not a valid range style capture");
  }

  SheetLayout& layout = sheet.mutable_layout();
  if ((flags & kCapturesColumns) != 0U) {
    set_column_style_span(layout, rect.first_col, rect.last_col, false, 0U);
    for (const ColumnSegment& seg : cols) {
      set_column_style_span(layout, seg.first, seg.last, true, seg.xf);
    }
  }
  if ((flags & kCapturesRows) != 0U) {
    RestoreRows ctx{&rows, 0U};
    rewrite_sparse_row_styles(
        layout, rect.first_row, rect.last_row, sheet.format_defaults().zero_height,
        [](void* raw, std::uint32_t row, RowDecision) {
          auto* restore = static_cast<RestoreRows*>(raw);
          if (restore->next < restore->rows->size() && (*restore->rows)[restore->next].row == row) {
            return RowDecision{true, (*restore->rows)[restore->next++].xf};
          }
          return RowDecision{};
        },
        &ctx, &rows);
  }
  // Clear every run and stored cell's own format, then put the captured ones back.
  const std::uint64_t col_end = static_cast<std::uint64_t>(rect.last_col) + 1U;
  for (const std::uint32_t row : rows_holding_styles(sheet, rect)) {
    sheet.clear_style_runs(row, rect.first_col, col_end);
  }
  for (const CapturedCell& cell : stored_cells(sheet, rect)) {
    sheet.clear_cell_xf_index(cell.row, cell.col);
  }
  for (const CapturedCell& cell : cells) {
    sheet.restore_cell_style(cell.row, cell.col, cell.xf, cell.explicit_xf);
  }
  for (const CapturedRun& run : runs) {
    std::vector<std::uint32_t> stored;
    if (const auto it = sheet.rows().find(run.row); it != sheet.rows().end()) {
      for (const RowCells::Entry entry : it->second.entries_in(run.first, static_cast<std::uint64_t>(run.last) + 1U)) {
        stored.push_back(entry.col);
      }
    }
    for (const std::uint32_t col : stored) {
      sheet.restore_cell_style(run.row, col, run.xf, true);
    }
    sheet.set_style_runs(run.row, run.first, static_cast<std::uint64_t>(run.last) + 1U, run.xf);
  }
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
