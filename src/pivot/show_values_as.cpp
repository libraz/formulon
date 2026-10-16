
#include "pivot/show_values_as.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "pivot/pivot_index.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using Path = std::vector<std::string>;

// One rendered cell as the transforms see it: blank, number or error.
struct Cell {
  enum class Kind { Blank, Num, Err };
  Kind kind = Kind::Blank;
  double n = 0.0;
  ErrorCode err = ErrorCode::Null;

  bool blank() const { return kind == Kind::Blank; }
  bool num() const { return kind == Kind::Num; }
  bool error() const { return kind == Kind::Err; }
};

Cell make_num(double n) {
  Cell c;
  c.kind = Cell::Kind::Num;
  c.n = n == 0.0 ? 0.0 : n;  // Excel never renders -0
  return c;
}

Cell make_err(ErrorCode e) {
  Cell c;
  c.kind = Cell::Kind::Err;
  c.err = e;
  return c;
}

Cell from_value(const Value& v) {
  if (v.is_number()) {
    return make_num(v.as_number());
  }
  if (v.is_error()) {
    return make_err(v.as_error());
  }
  return Cell{};
}

Value to_value(const Cell& c) {
  switch (c.kind) {
    case Cell::Kind::Num:
      return Value::number(c.n);
    case Cell::Kind::Err:
      return Value::error(c.err);
    case Cell::Kind::Blank:
      break;
  }
  return Value::blank();
}

// Blank reads as 0; numbers and errors pass through.
Cell zero_if_blank(const Cell& c) {
  return c.blank() ? make_num(0.0) : c;
}

// a / b: an error operand wins (a first), a zero or blank divisor is #DIV/0!.
Cell divide(const Cell& a, const Cell& b) {
  if (a.error()) {
    return a;
  }
  if (b.error()) {
    return b;
  }
  if (b.blank() || b.n == 0.0) {
    return make_err(ErrorCode::Div0);
  }
  return make_num(a.n / b.n);
}

struct Line {
  enum class Kind { Leaf, Sub, Grand };
  Kind kind = Kind::Leaf;
  std::size_t idx = 0;
  Path path;
  // Custom-function subtotal; addressed by index only, never by path.
  bool custom = false;
};

void collect_leaf_paths(const std::vector<AxisHierarchyNode>& nodes, Path& prefix, std::vector<Path>& out) {
  for (const AxisHierarchyNode& node : nodes) {
    prefix.push_back(node.label);
    if (node.children.empty()) {
      out.push_back(prefix);
    } else {
      collect_leaf_paths(node.children, prefix, out);
    }
    prefix.pop_back();
  }
}

// Leaves, subtotals and the grand total of one axis, addressable by item path.
struct Axis {
  std::vector<Line> lines;
  std::size_t grand = 0;

  std::optional<std::size_t> find(const Path& path) const {
    for (std::size_t i = 0; i < lines.size(); ++i) {
      if (!lines[i].custom && lines[i].path == path) {
        return i;
      }
    }
    return std::nullopt;
  }

  // Lines whose path equals `p` except at depth `d`, in display order.
  std::vector<std::size_t> siblings(const Path& p, std::size_t d) const {
    std::vector<std::size_t> out;
    for (std::size_t i = 0; i < lines.size(); ++i) {
      const Line& l = lines[i];
      if (l.custom || l.kind == Line::Kind::Grand || l.path.size() != p.size()) {
        continue;
      }
      bool same = true;
      for (std::size_t k = 0; k < p.size(); ++k) {
        same = same && (k == d || l.path[k] == p[k]);
      }
      if (same) {
        out.push_back(i);
      }
    }
    return out;
  }
};

template <typename Labels, typename Custom>
bool build_axis(const std::vector<AxisHierarchyNode>& tree, std::size_t leaf_count, std::size_t sub_count,
                const Labels& sub_labels, const Custom& sub_custom, Axis& axis) {
  std::vector<Path> leaves;
  Path prefix;
  collect_leaf_paths(tree, prefix, leaves);
  if (tree.empty()) {
    leaves.assign(leaf_count, Path{});
  }
  if (leaves.size() != leaf_count) {
    return false;
  }
  for (std::size_t i = 0; i < leaves.size(); ++i) {
    axis.lines.push_back({Line::Kind::Leaf, i, std::move(leaves[i]), false});
  }
  for (std::size_t i = 0; i < sub_count; ++i) {
    axis.lines.push_back({Line::Kind::Sub, i, sub_labels(i), sub_custom(i)});
  }
  axis.grand = axis.lines.size();
  axis.lines.push_back({Line::Kind::Grand, 0, Path{}, false});
  return true;
}

// Storage slot of data field `df` at the crossing of a row line and a column line.
Value* slot(PivotResult& res, std::size_t df, const Line& r, const Line& c) {
  const auto at = [df](std::vector<Value>& v) -> Value* { return df < v.size() ? &v[df] : nullptr; };
  using K = Line::Kind;
  switch (r.kind) {
    case K::Leaf:
      switch (c.kind) {
        case K::Leaf:
          return r.idx < res.values.size() && c.idx < res.values[r.idx].size() ? at(res.values[r.idx][c.idx]) : nullptr;
        case K::Sub:
          return c.idx < res.col_subtotals.size() && r.idx < res.col_subtotals[c.idx].values.size()
                     ? at(res.col_subtotals[c.idx].values[r.idx])
                     : nullptr;
        case K::Grand:
          return r.idx < res.row_leaf_totals.size() ? at(res.row_leaf_totals[r.idx]) : nullptr;
      }
      break;
    case K::Sub: {
      if (r.idx >= res.row_subtotals.size()) {
        return nullptr;
      }
      RowSubtotal& sub = res.row_subtotals[r.idx];
      switch (c.kind) {
        case K::Leaf:
          return c.idx < sub.col_values.size() ? at(sub.col_values[c.idx]) : nullptr;
        case K::Sub:
          return c.idx < sub.col_subtotal_values.size() ? at(sub.col_subtotal_values[c.idx]) : nullptr;
        case K::Grand:
          return at(sub.values);
      }
      break;
    }
    case K::Grand:
      switch (c.kind) {
        case K::Leaf:
          return c.idx < res.col_leaf_totals.size() ? at(res.col_leaf_totals[c.idx]) : nullptr;
        case K::Sub:
          return c.idx < res.col_subtotals.size() ? at(res.col_subtotals[c.idx].total) : nullptr;
        case K::Grand:
          return at(res.grand_totals);
      }
      break;
  }
  return nullptr;
}

struct BaseRef {
  bool on_row = true;
  std::size_t depth = 0;
};

std::optional<std::size_t> find_in(const std::vector<std::uint32_t>& order, std::uint32_t fi) {
  for (std::size_t i = 0; i < order.size(); ++i) {
    if (order[i] == fi) {
      return i;
    }
  }
  return std::nullopt;
}

// Axis and depth of the data field's base field. Without a base field a running total follows
// its mode's direction and the other base-field modes follow a lone single-level axis.
std::optional<BaseRef> resolve_base(const PivotTable& table, const PivotDataField& df, std::size_t row_leaves,
                                    std::size_t col_leaves) {
  if (df.show_as_base_field.has_value()) {
    if (const auto d = find_in(table.row_field_order(), *df.show_as_base_field); d.has_value()) {
      return BaseRef{true, *d};
    }
    if (const auto d = find_in(table.col_field_order(), *df.show_as_base_field); d.has_value()) {
      return BaseRef{false, *d};
    }
    return std::nullopt;
  }
  if (df.show_as == ShowValuesAs::RunningTotalInCol) {
    return BaseRef{true, 0};
  }
  if (df.show_as == ShowValuesAs::RunningTotalInRow) {
    return BaseRef{false, 0};
  }
  if (table.row_field_order().size() == 1 && row_leaves > 1) {
    return BaseRef{true, 0};
  }
  if (table.col_field_order().size() == 1 && col_leaves > 1) {
    return BaseRef{false, 0};
  }
  return std::nullopt;
}

// Everything one cell's transform reads: the snapshot of the untransformed grid and both axes.
struct Grid {
  const Axis& rows;
  const Axis& cols;
  std::vector<std::vector<Cell>> n;  // n[row line][col line]

  const Cell& at(std::size_t ri, std::size_t ci) const { return n[ri][ci]; }
  Cell grand() const { return n[rows.grand][cols.grand]; }
};

struct BaseSpec {
  ShowValuesAs mode;
  BaseRef base;
  std::optional<std::uint32_t> item;
  std::optional<std::string> item_label;  // resolved label of a specific baseItem
};

// Cell on the base axis at line `li`, other axis fixed at the current cell.
Cell base_cell(const Grid& g, const BaseSpec& s, std::size_t ri, std::size_t ci, std::size_t li) {
  return s.base.on_row ? g.at(li, ci) : g.at(ri, li);
}

Cell base_cell_by_path(const Grid& g, const BaseSpec& s, std::size_t ri, std::size_t ci, const Path& q) {
  const Axis& ax = s.base.on_row ? g.rows : g.cols;
  const std::optional<std::size_t> li = ax.find(q);
  return li.has_value() ? base_cell(g, s, ri, ci, *li) : Cell{};
}

Cell base_mode_cell(const Grid& g, const BaseSpec& s, std::size_t ri, std::size_t ci) {
  const Axis& ax = s.base.on_row ? g.rows : g.cols;
  const std::size_t cur = s.base.on_row ? ri : ci;
  const Line& line = ax.lines[cur];
  const Path& p = line.path;
  const std::size_t d = s.base.depth;
  const Cell c = g.at(ri, ci);
  const ShowValuesAs mode = s.mode;
  if (line.kind == Line::Kind::Grand || p.size() <= d) {
    if (mode == ShowValuesAs::PercentOfRunningTotal && c.error()) {
      return c;
    }
    return Cell{};
  }
  if (mode == ShowValuesAs::PercentOfParent) {
    const Path parent(p.begin(), p.begin() + static_cast<std::ptrdiff_t>(d) + 1);
    const Cell pv = base_cell_by_path(g, s, ri, ci, parent);
    if (pv.error() || pv.blank() || pv.n == 0.0) {
      return Cell{};
    }
    return c.error() ? c : make_num(zero_if_blank(c).n / pv.n);
  }
  const std::vector<std::size_t> sib = ax.siblings(p, d);
  const auto it = std::find(sib.begin(), sib.end(), cur);
  if (it == sib.end()) {
    return Cell{};
  }
  const std::size_t pos = static_cast<std::size_t>(it - sib.begin());

  if (mode == ShowValuesAs::RunningTotalInRow || mode == ShowValuesAs::RunningTotalInCol ||
      mode == ShowValuesAs::PercentOfRunningTotal) {
    Cell run;
    for (std::size_t k = 0; k <= pos; ++k) {
      const Cell x = base_cell(g, s, ri, ci, sib[k]);
      if (x.error()) {
        run = x;
        break;
      }
      if (!x.blank()) {
        run = run.blank() ? x : make_num(run.n + x.n);
      }
    }
    if (run.blank()) {
      run = make_num(0.0);
    }
    if (mode != ShowValuesAs::PercentOfRunningTotal) {
      return run;
    }
    if (run.error()) {
      return run;
    }
    // At the base level the total across the base field is the cell with that level collapsed;
    // deeper rows divide by the sum over all their siblings. A total with no value is blank.
    Cell total;
    if (p.size() == d + 1) {
      total = base_cell_by_path(g, s, ri, ci, Path(p.begin(), p.begin() + static_cast<std::ptrdiff_t>(d)));
    } else {
      for (const std::size_t li : sib) {
        const Cell x = base_cell(g, s, ri, ci, li);
        if (x.error()) {
          total = x;
          break;
        }
        if (!x.blank()) {
          total = total.blank() ? x : make_num(total.n + x.n);
        }
      }
    }
    if (total.blank()) {
      return Cell{};
    }
    if (total.error() || total.n == 0.0) {
      return make_err(ErrorCode::Div0);
    }
    return make_num(run.n / total.n);
  }

  if (mode == ShowValuesAs::RankAscending || mode == ShowValuesAs::RankDescending) {
    if (!c.num()) {
      return Cell{};
    }
    std::size_t rank = 1;
    std::vector<double> seen;
    for (const std::size_t li : sib) {
      const Cell x = base_cell(g, s, ri, ci, li);
      if (!x.num() || std::find(seen.begin(), seen.end(), x.n) != seen.end()) {
        continue;
      }
      seen.push_back(x.n);
      const bool before = mode == ShowValuesAs::RankAscending ? x.n < c.n : x.n > c.n;
      rank += before ? 1U : 0U;
    }
    return make_num(static_cast<double>(rank));
  }

  // Difference / percent difference / percent of a base item.
  std::optional<std::size_t> bq;
  if (s.item == kShowAsBasePrev) {
    bq = pos > 0 ? sib[pos - 1] : cur;
  } else if (s.item == kShowAsBaseNext) {
    bq = pos + 1 < sib.size() ? sib[pos + 1] : cur;
  } else if (s.item_label.has_value()) {
    for (const std::size_t li : sib) {
      if (ax.lines[li].path[d] == *s.item_label) {
        bq = li;
        break;
      }
    }
  }
  const Cell b = bq.has_value() ? base_cell(g, s, ri, ci, *bq) : Cell{};
  const bool self = bq.has_value() && *bq == cur;
  if (mode == ShowValuesAs::DifferenceFrom) {
    if (self) {
      return Cell{};
    }
    if (c.error()) {
      return c;
    }
    if (b.error()) {
      return b;
    }
    return make_num(zero_if_blank(c).n - zero_if_blank(b).n);
  }
  if (self && mode == ShowValuesAs::PercentDifferenceFrom) {
    return Cell{};
  }
  if (c.error()) {
    return c;
  }
  if (b.blank()) {
    return Cell{};
  }
  if (c.blank()) {
    return make_err(ErrorCode::Null);
  }
  if (b.error()) {
    return Cell{};
  }
  if (b.n == 0.0) {
    return make_err(ErrorCode::Div0);
  }
  return make_num(mode == ShowValuesAs::PercentDifferenceFrom ? (c.n - b.n) / b.n : c.n / b.n);
}

Cell parent_axis_cell(const Grid& g, ShowValuesAs mode, std::size_t ri, std::size_t ci) {
  const bool on_row = mode == ShowValuesAs::PercentOfParentRow;
  const Axis& ax = on_row ? g.rows : g.cols;
  const std::size_t cur = on_row ? ri : ci;
  const Path& p = ax.lines[cur].path;
  // A grand total is its own parent; so is an item of an axis with no fields.
  std::optional<std::size_t> parent = cur;
  if (!p.empty()) {
    parent = ax.find(Path(p.begin(), p.end() - 1));
  }
  if (!parent.has_value()) {
    return Cell{};
  }
  const Cell c = g.at(ri, ci);
  const Cell pv = on_row ? g.at(*parent, ci) : g.at(ri, *parent);
  if (pv.error() || pv.blank() || pv.n == 0.0) {
    return Cell{};
  }
  return c.error() ? c : make_num(zero_if_blank(c).n / pv.n);
}

Cell transform_cell(const Grid& g, const BaseSpec* base, ShowValuesAs mode, std::size_t ri, std::size_t ci) {
  const Cell c = g.at(ri, ci);
  switch (mode) {
    case ShowValuesAs::PercentOfTotal:
      return divide(zero_if_blank(c), g.grand());
    case ShowValuesAs::PercentOfRow:
      return divide(zero_if_blank(c), g.at(ri, g.cols.grand));
    case ShowValuesAs::PercentOfCol:
      return divide(zero_if_blank(c), g.at(g.rows.grand, ci));
    case ShowValuesAs::Index: {
      const Cell x = zero_if_blank(c);
      const Cell gt = g.grand();
      const Cell rt = g.at(ri, g.cols.grand);
      const Cell ct = g.at(g.rows.grand, ci);
      for (const Cell* e : {&x, &gt, &rt, &ct}) {
        if (e->error()) {
          return *e;
        }
      }
      if (rt.blank() || ct.blank() || rt.n == 0.0 || ct.n == 0.0) {
        return make_err(ErrorCode::Div0);
      }
      return make_num(((x.n * gt.n) / rt.n) / ct.n);
    }
    case ShowValuesAs::PercentOfParentRow:
    case ShowValuesAs::PercentOfParentCol:
      return parent_axis_cell(g, mode, ri, ci);
    case ShowValuesAs::Normal:
      return c;
    default:
      return base != nullptr ? base_mode_cell(g, *base, ri, ci) : c;
  }
}

bool is_base_mode(ShowValuesAs mode) {
  switch (mode) {
    case ShowValuesAs::RunningTotalInRow:
    case ShowValuesAs::RunningTotalInCol:
    case ShowValuesAs::DifferenceFrom:
    case ShowValuesAs::PercentDifferenceFrom:
    case ShowValuesAs::PercentOf:
    case ShowValuesAs::PercentOfRunningTotal:
    case ShowValuesAs::PercentOfParent:
    case ShowValuesAs::RankAscending:
    case ShowValuesAs::RankDescending:
      return true;
    default:
      return false;
  }
}

// Keeps the legacy mirrors in step: the flat `subtotals` view tracks `row_subtotals[i].values`
// and `grand_total` tracks `grand_totals[0]`.
void sync_legacy_mirrors(PivotResult& result) {
  for (std::size_t i = 0; i < result.row_subtotals.size() && i < result.subtotals.size(); ++i) {
    const std::size_t n = std::min(result.subtotals[i].size(), result.row_subtotals[i].values.size());
    for (std::size_t df_idx = 0; df_idx < n; ++df_idx) {
      result.subtotals[i][df_idx] = result.row_subtotals[i].values[df_idx];
    }
  }
  if (!result.grand_totals.empty()) {
    result.grand_total = result.grand_totals[0];
  }
}

}  // namespace

static void apply_transforms(const PivotTable& table, const PivotCache& cache, PivotResult& result) {
  const std::size_t data_field_count = table.data_fields().size();
  if (result.values.empty() || data_field_count == 0) {
    return;
  }
  bool any = false;
  for (const PivotDataField& df : table.data_fields()) {
    any = any || df.show_as != ShowValuesAs::Normal;
  }
  if (!any) {
    return;
  }

  Axis rows;
  Axis cols;
  const bool built = build_axis(
                         result.rows, result.values.size(), result.row_subtotals.size(),
                         [&](std::size_t i) { return result.row_subtotals[i].labels; },
                         [&](std::size_t i) { return result.row_subtotals[i].aggregation.has_value(); }, rows) &&
                     build_axis(
                         result.cols, result.values[0].size(), result.col_subtotals.size(),
                         [&](std::size_t i) { return result.col_subtotals[i].labels; },
                         [&](std::size_t i) { return result.col_subtotals[i].aggregation.has_value(); }, cols);
  if (!built) {
    return;
  }
  const std::size_t row_leaves = result.values.size();
  const std::size_t col_leaves = result.values[0].size();

  for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
    const PivotDataField& df = table.data_fields()[df_idx];
    const ShowValuesAs mode = df.show_as;
    if (mode == ShowValuesAs::Normal) {
      continue;
    }
    std::optional<BaseSpec> base;
    if (is_base_mode(mode)) {
      const std::optional<BaseRef> ref = resolve_base(table, df, row_leaves, col_leaves);
      if (!ref.has_value()) {
        continue;
      }
      BaseSpec spec{mode, *ref, df.show_as_base_item, std::nullopt};
      if (!spec.item.has_value()) {
        spec.item = kShowAsBasePrev;
      }
      if (*spec.item != kShowAsBasePrev && *spec.item != kShowAsBaseNext && df.show_as_base_field.has_value() &&
          *df.show_as_base_field < table.fields().size()) {
        const auto& items = table.fields()[*df.show_as_base_field].items;
        if (*spec.item < items.size()) {
          spec.item_label = pivot_item_label(cache, *df.show_as_base_field, items[*spec.item]);
        }
      }
      base = spec;
    }

    Grid grid{rows, cols, {}};
    grid.n.assign(rows.lines.size(), std::vector<Cell>(cols.lines.size()));
    for (std::size_t ri = 0; ri < rows.lines.size(); ++ri) {
      for (std::size_t ci = 0; ci < cols.lines.size(); ++ci) {
        if (const Value* v = slot(result, df_idx, rows.lines[ri], cols.lines[ci]); v != nullptr) {
          grid.n[ri][ci] = from_value(*v);
        }
      }
    }
    // Grand totals switched off leave no slot: the total is still the sum of the leaf cells.
    if (slot(result, df_idx, rows.lines[rows.grand], cols.lines[cols.grand]) == nullptr) {
      Cell total;
      for (std::size_t r = 0; r < row_leaves; ++r) {
        for (std::size_t c = 0; c < col_leaves; ++c) {
          const Cell& x = grid.n[r][c];
          if (x.num()) {
            total = total.blank() ? x : make_num(total.n + x.n);
          }
        }
      }
      grid.n[rows.grand][cols.grand] = total;
    }

    for (std::size_t ri = 0; ri < rows.lines.size(); ++ri) {
      for (std::size_t ci = 0; ci < cols.lines.size(); ++ci) {
        Value* v = slot(result, df_idx, rows.lines[ri], cols.lines[ci]);
        if (v == nullptr) {
          continue;
        }
        *v = to_value(transform_cell(grid, base.has_value() ? &*base : nullptr, mode, ri, ci));
      }
    }
  }
}

void apply_show_values_as_transforms(const PivotTable& table, const PivotCache& cache, PivotResult& result,
                                     const std::vector<std::vector<std::size_t>>& /*row_subtotal_leaf_sets*/,
                                     const std::vector<std::vector<std::size_t>>& /*col_subtotal_leaf_sets*/) {
  apply_transforms(table, cache, result);
  sync_legacy_mirrors(result);
}

}  // namespace formulon::pivot
