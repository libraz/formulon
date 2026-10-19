#include "style_region.h"

#include <algorithm>
#include <utility>

#include "cell.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "style_resolve.h"
#include "utils/index_sort.h"

namespace formulon {
namespace {

constexpr std::uint8_t kSnapshotMagic[4] = {'F', 'M', 'R', 'S'};
constexpr std::uint8_t kSnapshotVersion = 1U;
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

std::uint32_t apply_map(const std::vector<XfRemap>& map, std::uint32_t xf) {
  const auto it = std::lower_bound(map.begin(), map.end(), xf,
                                   [](const XfRemap& pair, std::uint32_t key) { return pair.from < key; });
  return (it != map.end() && it->from == xf) ? it->to : xf;
}

// Every stored cell of `rect` with its style state; `explicit_xf` here means
// the cell owns its style.
std::vector<CapturedCell> stored_cells(const Sheet& sheet, const StyleRect& rect) {
  std::vector<CapturedCell> out;
  for (const auto& [row, cells] : sheet.rows()) {
    if (row < rect.first_row || row > rect.last_row || cells.empty()) {
      continue;
    }
    for (const auto& [col, cell] : cells.entries_in(rect.first_col, static_cast<std::uint64_t>(rect.last_col) + 1U)) {
      out.push_back({row, col, cell.xf_index, cell_owns_style(cell)});
    }
  }
  return out;
}

// Decides the style of one row of a rewrite from its current style; a result
// with `styled == false` leaves the row unstyled.
struct RowDecision {
  bool styled = false;
  std::uint32_t xf = 0;
};
using RowDecide = RowDecision (*)(void* ctx, std::uint32_t row, RowDecision current);

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
    const RowDecision decision = decide(ctx, entry.row, {entry.has_style, entry.style_xf});
    entry.has_style = decision.styled;
    entry.style_xf = decision.styled ? decision.xf : 0U;
    if (!should_erase_row_override(entry, default_hidden, existed_before)) {
      out.push_back(entry);
    }
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
    const RowDecision decision = decide(ctx, row, {entry.has_style, entry.style_xf});
    entry.has_style = decision.styled;
    entry.style_xf = decision.styled ? decision.xf : 0U;
    if (!should_erase_row_override(entry, default_hidden, existed_before)) {
      out.push_back(entry);
    }
  }
  out.insert(out.end(), rows.begin() + static_cast<std::ptrdiff_t>(next), rows.end());
  rows = std::move(out);
}

void append_u32(std::vector<std::uint8_t>& out, std::uint32_t value) {
  for (int shift = 0; shift < 32; shift += 8) {
    out.push_back(static_cast<std::uint8_t>((value >> shift) & 0xFFU));
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
  const std::vector<XfRemap>* map;
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

std::vector<std::uint32_t> range_style_bases(const Sheet& sheet, const StyleRect& rect) {
  const std::vector<RowStyle> rows = styled_rows(sheet.layout(), rect.first_row, rect.last_row);
  const std::vector<ColumnSegment> cols = column_segments(sheet.layout(), rect.first_col, rect.last_col);
  // Owning cells per styled row, and per column segment across the unstyled rows.
  std::vector<std::uint64_t> owned_in_row(rows.size(), 0U);
  std::vector<std::uint64_t> owned_in_col(cols.size(), 0U);
  std::vector<std::uint32_t> out;
  for (const CapturedCell& cell : stored_cells(sheet, rect)) {
    if (!cell.explicit_xf) {
      continue;
    }
    out.push_back(cell.xf);
    const auto row_it = std::lower_bound(rows.begin(), rows.end(), cell.row,
                                         [](const RowStyle& row, std::uint32_t key) { return row.row < key; });
    if (row_it != rows.end() && row_it->row == cell.row) {
      ++owned_in_row[static_cast<std::size_t>(row_it - rows.begin())];
      continue;
    }
    const auto col_it = std::upper_bound(cols.begin(), cols.end(), cell.col,
                                         [](std::uint32_t key, const ColumnSegment& seg) { return key < seg.first; });
    ++owned_in_col[static_cast<std::size_t>(col_it - cols.begin()) - 1U];
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

Expected<void, Error> remap_range_style(Sheet& sheet, const StyleRect& rect, const std::vector<XfRemap>& map,
                                        std::uint64_t cell_limit) {
  const bool all_rows = covers_all_rows(rect);
  const bool all_cols = covers_all_cols(rect);
  const std::uint64_t area = (static_cast<std::uint64_t>(rect.last_row - rect.first_row) + 1U) *
                             (static_cast<std::uint64_t>(rect.last_col - rect.first_col) + 1U);
  if (!all_rows && !all_cols && area > cell_limit) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "remap_range_style: too many cells to restyle");
  }
  std::vector<CapturedCell> owned = stored_cells(sheet, rect);
  std::uint64_t written = static_cast<std::uint64_t>(
      std::count_if(owned.begin(), owned.end(), [](const CapturedCell& cell) { return cell.explicit_xf; }));
  if (written > cell_limit) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "remap_range_style: too many cells to restyle");
  }

  const bool default_hidden = sheet.format_defaults().zero_height;
  const std::uint32_t mapped_default = apply_map(map, 0U);
  const SheetLayout& layout = sheet.layout();
  const std::vector<ColumnSegment> pre_cols = column_segments(layout, rect.first_col, rect.last_col);
  const std::vector<RowStyle> pre_rows = styled_rows(layout, rect.first_row, rect.last_row);

  std::vector<ColumnSegment> post_cols = pre_cols;
  if (all_rows) {
    for (ColumnSegment& seg : post_cols) {
      seg.xf = apply_map(map, seg.styled ? seg.xf : 0U);
      seg.styled = seg.styled || mapped_default != 0U;
    }
  }
  const bool style_bare_rows = all_cols && !all_rows && mapped_default != 0U;
  // A row band leaves column styles in place, so bare rows under a remapped column need cell writes.
  const bool bare_rows_cross_remapped_cols =
      all_cols && !all_rows && std::any_of(pre_cols.begin(), pre_cols.end(), [&map](const ColumnSegment& seg) {
        return seg.styled && apply_map(map, seg.xf) != seg.xf;
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
    const RowCells* stored_row = nullptr;
    bool looked_up_stored_row = false;
    std::uint32_t post_row_xf = 0U;
    bool post_row_styled = pre_row != nullptr;
    if (pre_row != nullptr) {
      post_row_xf = all_cols ? apply_map(map, pre_row->xf) : pre_row->xf;
    } else if (style_bare_rows) {
      post_row_styled = true;
      post_row_xf = mapped_default;
    }
    for (std::size_t i = 0; i < pre_cols.size(); ++i) {
      const ColumnSegment& pre = pre_cols[i];
      const ColumnSegment& post = post_cols[i];
      const std::uint32_t target = apply_map(map, pre_row != nullptr ? pre_row->xf : (pre.styled ? pre.xf : 0U));
      const std::uint32_t after = post_row_styled ? post_row_xf : (post.styled ? post.xf : 0U);
      if ((all_rows || all_cols) && after == target) {
        continue;
      }
      runs.push_back({row, pre.first, pre.last, target});
      if (!looked_up_stored_row) {
        const auto stored_row_it = sheet.rows().find(row);
        stored_row = stored_row_it != sheet.rows().end() ? &stored_row_it->second : nullptr;
        looked_up_stored_row = true;
      }
      const std::uint64_t run_width = static_cast<std::uint64_t>(pre.last - pre.first) + 1U;
      std::uint64_t owned_in_run = 0U;
      if (stored_row != nullptr) {
        for (const RowCells::Entry entry :
             stored_row->entries_in(pre.first, static_cast<std::uint64_t>(pre.last) + 1U)) {
          if (cell_owns_style(entry.cell)) {
            ++owned_in_run;
          }
        }
      }
      // Owner cells are written once by the owner pass below, unchanged xf included.
      written += run_width - owned_in_run;
    }
  };
  if ((all_rows || all_cols) && !style_bare_rows && !bare_rows_cross_remapped_cols) {
    // Only styled row overrides can require cell writes.
    for (std::size_t i = 0; i < pre_rows.size() && written <= cell_limit; ++i) {
      plan_row(pre_rows[i].row, &pre_rows[i]);
    }
  } else {
    std::size_t next = 0;
    for (std::uint64_t row = rect.first_row; row <= rect.last_row && written <= cell_limit; ++row) {
      const RowStyle* pre_row = nullptr;
      if (next < pre_rows.size() && pre_rows[next].row == row) {
        pre_row = &pre_rows[next++];
      }
      plan_row(static_cast<std::uint32_t>(row), pre_row);
    }
  }
  if (written > cell_limit) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "remap_range_style: too many cells to restyle");
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
        return RowDecision{true, apply_map(*remap->map, current.xf)};
      }
      return RowDecision{remap->style_bare, remap->mapped_default};
    };
    if (style_bare_rows) {
      rewrite_row_styles(mutable_layout, rect.first_row, rect.last_row, default_hidden, remap_row, &ctx);
    } else {
      rewrite_sparse_row_styles(mutable_layout, rect.first_row, rect.last_row, default_hidden, remap_row, &ctx);
    }
  }
  for (const CellRun& run : runs) {
    for (std::uint64_t col = run.first; col <= run.last; ++col) {
      const Cell* cell = sheet.cell_at(run.row, static_cast<std::uint32_t>(col));
      if (cell == nullptr || !cell_owns_style(*cell)) {
        sheet.set_cell_xf_index(run.row, static_cast<std::uint32_t>(col), run.xf);
      }
    }
  }
  for (const CapturedCell& cell : owned) {
    if (cell.explicit_xf) {
      sheet.set_cell_xf_index(cell.row, cell.col, apply_map(map, cell.xf));
    }
  }
  return Expected<void, Error>::Ok();
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
  std::vector<CapturedCell> cells;
  for (const CapturedCell& cell : stored_cells(sheet, rect)) {
    const Cell* stored = sheet.cell_at(cell.row, cell.col);
    if (stored != nullptr && (stored->has_explicit_xf || stored->xf_index != 0U)) {
      cells.push_back({cell.row, cell.col, stored->xf_index, stored->has_explicit_xf});
    }
  }
  append_u32(out, static_cast<std::uint32_t>(cells.size()));
  for (const CapturedCell& cell : cells) {
    append_u32(out, cell.row);
    append_u32(out, cell.col);
    append_u32(out, cell.xf);
    out.push_back(cell.explicit_xf ? 1U : 0U);
  }
  return out;
}

namespace {

// Decodes a `snapshot_range_style` capture; false on any malformed field.
bool parse_capture(const std::uint8_t* bytes, std::size_t len, std::size_t xf_count, StyleRect& rect,
                   std::uint8_t& flags, std::vector<ColumnSegment>& cols, std::vector<RowStyle>& rows,
                   std::vector<CapturedCell>& cells) {
  if (bytes == nullptr || len < sizeof(kSnapshotMagic) ||
      !std::equal(std::begin(kSnapshotMagic), std::end(kSnapshotMagic), bytes)) {
    return false;
  }
  ByteReader in(bytes + sizeof(kSnapshotMagic), len - sizeof(kSnapshotMagic));
  std::uint8_t version = 0;
  if (!in.u8(version) || version != kSnapshotVersion || !in.u32(rect.first_row) || !in.u32(rect.first_col) ||
      !in.u32(rect.last_row) || !in.u32(rect.last_col) || !in.u8(flags) || rect.first_row > rect.last_row ||
      rect.first_col > rect.last_col || rect.last_row >= Sheet::kMaxRows || rect.last_col >= Sheet::kMaxCols ||
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
  return in.at_end();
}

}  // namespace

Expected<void, Error> restore_range_style(Sheet& sheet, const std::uint8_t* bytes, std::size_t len,
                                          std::size_t xf_count) {
  StyleRect rect;
  std::uint8_t flags = 0;
  std::vector<ColumnSegment> cols;
  std::vector<RowStyle> rows;
  std::vector<CapturedCell> cells;
  if (!parse_capture(bytes, len, xf_count, rect, flags, cols, rows, cells)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "restore_range_style: not a valid range style capture");
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
  // Clear every stored cell, then put the captured ones back.
  for (const CapturedCell& cell : stored_cells(sheet, rect)) {
    sheet.clear_cell_xf_index(cell.row, cell.col);
  }
  for (const CapturedCell& cell : cells) {
    sheet.restore_cell_style(cell.row, cell.col, cell.xf, cell.explicit_xf);
  }
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
