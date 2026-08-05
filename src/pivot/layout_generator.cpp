
#include "pivot/layout_generator.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <utility>
#include <vector>

#include "pivot/aggregator.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::pivot {

void apply_show_values_as_transforms(const PivotTable& table, PivotResult& result,
                                     const std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                                     const std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets) {
  const std::size_t data_field_count = table.data_fields().size();
  if (result.values.empty() || data_field_count == 0) {
    return;
  }

  auto cell_num = [](const Value& v) -> std::pair<bool, double> {
    if (v.is_number()) {
      return {true, v.as_number()};
    }
    if (v.is_boolean()) {
      return {true, v.as_boolean() ? 1.0 : 0.0};
    }
    return {false, 0.0};
  };
  // Scales `cell` in place by `denom`. Only acts when `cell` is a
  // numeric aggregate; non-numeric (blank / text / error) slots are
  // left untouched. Emits Div0 when `denom == 0` and the slot was
  // numeric.
  auto scale_cell = [&cell_num](Value& cell, double denom) {
    auto [ok, n] = cell_num(cell);
    if (!ok) {
      return;
    }
    if (denom == 0.0) {
      cell = Value::error(ErrorCode::Div0);
    } else {
      cell = Value::number(n / denom);
    }
  };
  const std::size_t actual_row_count = result.values.size();
  const std::size_t actual_col_count = result.values[0].size();

  for (std::size_t df_idx = 0; df_idx < data_field_count; ++df_idx) {
    const ShowValuesAs mode = table.data_fields()[df_idx].show_as;
    if (mode == ShowValuesAs::Normal) {
      continue;
    }
    switch (mode) {
      case ShowValuesAs::PercentOfRow: {
        // Per-leaf-row sums (used for both leaf cells and any
        // col_subtotal cell that lives in that leaf row).
        std::vector<double> row_sums(actual_row_count, 0.0);
        std::vector<bool> row_any_numeric(actual_row_count, false);
        for (std::size_t r = 0; r < actual_row_count; ++r) {
          for (std::size_t c = 0; c < actual_col_count; ++c) {
            if (df_idx >= result.values[r][c].size()) {
              continue;
            }
            auto [ok, n] = cell_num(result.values[r][c][df_idx]);
            if (ok) {
              row_sums[r] += n;
              row_any_numeric[r] = true;
            }
          }
        }
        for (std::size_t r = 0; r < actual_row_count; ++r) {
          for (std::size_t c = 0; c < actual_col_count; ++c) {
            if (df_idx >= result.values[r][c].size()) {
              continue;
            }
            if (!row_any_numeric[r]) {
              continue;
            }
            scale_cell(result.values[r][c][df_idx], row_sums[r]);
          }
        }
        // Row subtotals: each is its own "row" with its own row sum
        // taken over `col_values`. The total slot becomes 1.0 (all of
        // the row's contribution lives within itself); col_values are
        // their share of that sum; col_subtotal_values are partial
        // shares of the same row sum, matching Excel's "% of row"
        // treatment of intersection cells.
        for (RowSubtotal& sub : result.row_subtotals) {
          double sub_row_sum = 0.0;
          bool sub_row_any_numeric = false;
          for (const auto& col_slot : sub.col_values) {
            if (df_idx >= col_slot.size()) {
              continue;
            }
            auto [ok, n] = cell_num(col_slot[df_idx]);
            if (ok) {
              sub_row_sum += n;
              sub_row_any_numeric = true;
            }
          }
          for (auto& col_slot : sub.col_values) {
            if (df_idx >= col_slot.size()) {
              continue;
            }
            if (!sub_row_any_numeric) {
              continue;
            }
            scale_cell(col_slot[df_idx], sub_row_sum);
          }
          for (auto& cs_slot : sub.col_subtotal_values) {
            if (df_idx >= cs_slot.size()) {
              continue;
            }
            if (!sub_row_any_numeric) {
              continue;
            }
            scale_cell(cs_slot[df_idx], sub_row_sum);
          }
          if (df_idx < sub.values.size() && sub_row_any_numeric) {
            if (sub_row_sum == 0.0) {
              sub.values[df_idx] = Value::error(ErrorCode::Div0);
            } else {
              sub.values[df_idx] = Value::number(1.0);
            }
          }
        }
        // Col subtotals: each cell sits in some leaf row `r`, so
        // divide by the same `row_sums[r]` used for the leaf-row
        // transform.
        for (ColSubtotal& csub : result.col_subtotals) {
          for (std::size_t r = 0; r < csub.values.size() && r < actual_row_count; ++r) {
            if (df_idx >= csub.values[r].size()) {
              continue;
            }
            if (!row_any_numeric[r]) {
              continue;
            }
            scale_cell(csub.values[r][df_idx], row_sums[r]);
          }
        }
        // Grand total: under PercentOfRow the grand-total row sums to
        // itself, so the displayed value is 1.0 (Div0 if no row had
        // any numeric content).
        if (df_idx < result.grand_totals.size()) {
          bool any_row_numeric = false;
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            if (row_any_numeric[r]) {
              any_row_numeric = true;
              break;
            }
          }
          auto [ok, _n] = cell_num(result.grand_totals[df_idx]);
          (void)_n;
          if (ok) {
            result.grand_totals[df_idx] = any_row_numeric ? Value::number(1.0) : Value::error(ErrorCode::Div0);
          }
        }
        break;
      }
      case ShowValuesAs::PercentOfCol: {
        // Per-leaf-col sums (mirror of PercentOfRow).
        std::vector<double> col_sums(actual_col_count, 0.0);
        std::vector<bool> col_any_numeric(actual_col_count, false);
        for (std::size_t c = 0; c < actual_col_count; ++c) {
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            if (c >= result.values[r].size() || df_idx >= result.values[r][c].size()) {
              continue;
            }
            auto [ok, n] = cell_num(result.values[r][c][df_idx]);
            if (ok) {
              col_sums[c] += n;
              col_any_numeric[c] = true;
            }
          }
        }
        // Capture the grand total before any mutation; under
        // PercentOfCol the row-subtotal "row total" slot collapses to
        // its share of the grand total (the row's contribution to the
        // single-column world that PercentOfCol presents).
        double total = 0.0;
        bool total_known = false;
        if (df_idx < result.grand_totals.size()) {
          auto [ok, n] = cell_num(result.grand_totals[df_idx]);
          if (ok) {
            total = n;
            total_known = true;
          }
        }
        if (!total_known) {
          for (std::size_t c = 0; c < actual_col_count; ++c) {
            total += col_sums[c];
            if (col_any_numeric[c]) {
              total_known = true;
            }
          }
        }
        // Col-subtotal column totals: per col_subtotal, the
        // subtotal-column total = sum across its `values[r][df_idx]`
        // slots. Used for both the col_subtotal cells themselves and
        // for any row_subtotal cell that lives in that col_subtotal.
        std::vector<double> col_subtotal_totals(result.col_subtotals.size(), 0.0);
        std::vector<bool> col_subtotal_any_numeric(result.col_subtotals.size(), false);
        for (std::size_t cs = 0; cs < result.col_subtotals.size(); ++cs) {
          for (const auto& row_slot : result.col_subtotals[cs].values) {
            if (df_idx >= row_slot.size()) {
              continue;
            }
            auto [ok, n] = cell_num(row_slot[df_idx]);
            if (ok) {
              col_subtotal_totals[cs] += n;
              col_subtotal_any_numeric[cs] = true;
            }
          }
        }
        for (std::size_t c = 0; c < actual_col_count; ++c) {
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            if (c >= result.values[r].size() || df_idx >= result.values[r][c].size()) {
              continue;
            }
            if (!col_any_numeric[c]) {
              continue;
            }
            scale_cell(result.values[r][c][df_idx], col_sums[c]);
          }
        }
        // Col subtotals: each col_subtotal column's cells divide by
        // that subtotal column's own sum (same shape as a leaf col).
        for (std::size_t cs = 0; cs < result.col_subtotals.size(); ++cs) {
          ColSubtotal& csub = result.col_subtotals[cs];
          for (auto& row_slot : csub.values) {
            if (df_idx >= row_slot.size()) {
              continue;
            }
            if (!col_subtotal_any_numeric[cs]) {
              continue;
            }
            scale_cell(row_slot[df_idx], col_subtotal_totals[cs]);
          }
        }
        // Row subtotals: each col_values[col] cell divides by that
        // leaf col's sum. Each col_subtotal_values[cs] cell divides by
        // the corresponding col-subtotal column total. The row's
        // overall `values[df]` slot collapses to its share of the
        // grand total (the row's contribution under col percentages).
        for (RowSubtotal& sub : result.row_subtotals) {
          for (std::size_t c = 0; c < sub.col_values.size() && c < actual_col_count; ++c) {
            if (df_idx >= sub.col_values[c].size()) {
              continue;
            }
            if (!col_any_numeric[c]) {
              continue;
            }
            scale_cell(sub.col_values[c][df_idx], col_sums[c]);
          }
          for (std::size_t cs = 0; cs < sub.col_subtotal_values.size() && cs < result.col_subtotals.size(); ++cs) {
            if (df_idx >= sub.col_subtotal_values[cs].size()) {
              continue;
            }
            if (!col_subtotal_any_numeric[cs]) {
              continue;
            }
            scale_cell(sub.col_subtotal_values[cs][df_idx], col_subtotal_totals[cs]);
          }
          if (df_idx < sub.values.size() && total_known) {
            scale_cell(sub.values[df_idx], total);
          }
        }
        // Grand total: under PercentOfCol every column sums to itself,
        // so the column-direction grand-total cell is 1.0.
        if (df_idx < result.grand_totals.size()) {
          bool any_col_numeric = false;
          for (std::size_t c = 0; c < actual_col_count; ++c) {
            if (col_any_numeric[c]) {
              any_col_numeric = true;
              break;
            }
          }
          auto [ok, _n] = cell_num(result.grand_totals[df_idx]);
          (void)_n;
          if (ok) {
            result.grand_totals[df_idx] = any_col_numeric ? Value::number(1.0) : Value::error(ErrorCode::Div0);
          }
        }
        break;
      }
      case ShowValuesAs::PercentOfTotal: {
        // Capture grand total once before mutating anything; reuse
        // the same denominator for every slot.
        double total = 0.0;
        bool total_known = false;
        if (df_idx < result.grand_totals.size()) {
          auto [ok, n] = cell_num(result.grand_totals[df_idx]);
          if (ok) {
            total = n;
            total_known = true;
          }
        }
        if (!total_known) {
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
              if (df_idx >= result.values[r][c].size()) {
                continue;
              }
              auto [ok, n] = cell_num(result.values[r][c][df_idx]);
              if (ok) {
                total += n;
                total_known = true;
              }
            }
          }
        }
        if (!total_known) {
          break;
        }
        for (std::size_t r = 0; r < actual_row_count; ++r) {
          for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
            if (df_idx >= result.values[r][c].size()) {
              continue;
            }
            scale_cell(result.values[r][c][df_idx], total);
          }
        }
        for (RowSubtotal& sub : result.row_subtotals) {
          if (df_idx < sub.values.size()) {
            scale_cell(sub.values[df_idx], total);
          }
          for (auto& col_slot : sub.col_values) {
            if (df_idx < col_slot.size()) {
              scale_cell(col_slot[df_idx], total);
            }
          }
          for (auto& cs_slot : sub.col_subtotal_values) {
            if (df_idx < cs_slot.size()) {
              scale_cell(cs_slot[df_idx], total);
            }
          }
        }
        for (ColSubtotal& csub : result.col_subtotals) {
          for (auto& row_slot : csub.values) {
            if (df_idx < row_slot.size()) {
              scale_cell(row_slot[df_idx], total);
            }
          }
        }
        if (df_idx < result.grand_totals.size()) {
          auto [ok, _n] = cell_num(result.grand_totals[df_idx]);
          (void)_n;
          if (ok) {
            result.grand_totals[df_idx] = (total == 0.0) ? Value::error(ErrorCode::Div0) : Value::number(1.0);
          }
        }
        break;
      }
      case ShowValuesAs::RunningTotalInRow:
      case ShowValuesAs::RunningTotalInCol: {
        // OOXML has a single `runTotal` show-data-as; the accumulation
        // direction is carried by the data field's `baseField`, not by the
        // attribute name. Resolve the direction from that base field's
        // axis, defaulting to the enum's baked direction when no base field
        // is set. Subtotals and grand totals are intentionally left at
        // their raw aggregate; see the header comment for this section.
        const PivotDataField& df = table.data_fields()[df_idx];
        bool run_in_col = (mode == ShowValuesAs::RunningTotalInCol);
        if (df.show_as_base_field.has_value()) {
          const std::uint32_t bf = *df.show_as_base_field;
          for (const std::uint32_t fi : table.col_field_order()) {
            if (fi == bf) {
              run_in_col = true;
              break;
            }
          }
          for (const std::uint32_t fi : table.row_field_order()) {
            if (fi == bf) {
              run_in_col = false;
              break;
            }
          }
        }
        if (run_in_col) {
          for (std::size_t c = 0; c < actual_col_count; ++c) {
            double running = 0.0;
            for (std::size_t r = 0; r < actual_row_count; ++r) {
              if (c >= result.values[r].size() || df_idx >= result.values[r][c].size()) {
                continue;
              }
              Value& cell = result.values[r][c][df_idx];
              auto [ok, n] = cell_num(cell);
              if (!ok) {
                continue;
              }
              running += n;
              cell = Value::number(running);
            }
          }
        } else {
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            double running = 0.0;
            for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
              if (df_idx >= result.values[r][c].size()) {
                continue;
              }
              Value& cell = result.values[r][c][df_idx];
              auto [ok, n] = cell_num(cell);
              if (!ok) {
                continue;
              }
              running += n;
              cell = Value::number(running);
            }
          }
        }
        break;
      }
      case ShowValuesAs::Index: {
        // Index = (cell * grand_total) / (row_sum * col_sum). Compute
        // partials on demand; if any partial is zero or non-numeric,
        // surface Div0 / leave as-is. Subtotals + grand totals remain
        // at their raw aggregate; see the header comment for this
        // section.
        double total = 0.0;
        bool total_known = false;
        if (df_idx < result.grand_totals.size()) {
          auto [ok, n] = cell_num(result.grand_totals[df_idx]);
          if (ok) {
            total = n;
            total_known = true;
          }
        }
        // Precompute row sums + col sums for this df.
        std::vector<double> row_sums(actual_row_count, 0.0);
        std::vector<double> col_sums(actual_col_count, 0.0);
        for (std::size_t r = 0; r < actual_row_count; ++r) {
          for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
            if (df_idx >= result.values[r][c].size()) {
              continue;
            }
            auto [ok, n] = cell_num(result.values[r][c][df_idx]);
            if (ok) {
              row_sums[r] += n;
              col_sums[c] += n;
            }
          }
        }
        // When grand totals are turned off the grand-total slot is empty,
        // which would leave `total == 0` and collapse every Index cell to
        // zero. Recompute the total from the surviving leaf cells, mirroring
        // the PercentOfTotal fallback.
        if (!total_known) {
          for (double rs : row_sums) {
            total += rs;
          }
        }
        for (std::size_t r = 0; r < actual_row_count; ++r) {
          for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
            if (df_idx >= result.values[r][c].size()) {
              continue;
            }
            Value& cell = result.values[r][c][df_idx];
            auto [ok, n] = cell_num(cell);
            if (!ok) {
              continue;
            }
            const double denom = row_sums[r] * col_sums[c];
            if (denom == 0.0) {
              cell = Value::error(ErrorCode::Div0);
            } else {
              cell = Value::number((n * total) / denom);
            }
          }
        }
        break;
      }
      case ShowValuesAs::DifferenceFrom:
      case ShowValuesAs::PercentDifferenceFrom: {
        // Resolve the base axis from `show_as_base_field`. We support
        // single-level base axis only: a base field that lives in a
        // multi-level hierarchy (e.g. `{Region, Product}` row order)
        // is treated as a no-op so the rendered values stay at their
        // raw aggregate. The same fallback applies when `base_field`
        // is unset and both axes are multi-level — the unambiguous
        // "previous" only makes sense when one axis has a single
        // ordering.
        const PivotDataField& df = table.data_fields()[df_idx];
        enum class BaseAxis { None, Row, Col } base_axis = BaseAxis::None;
        if (df.show_as_base_field.has_value()) {
          const std::uint32_t bf = *df.show_as_base_field;
          for (const std::uint32_t fi : table.row_field_order()) {
            if (fi == bf) {
              base_axis = BaseAxis::Row;
              break;
            }
          }
          if (base_axis == BaseAxis::None) {
            for (const std::uint32_t fi : table.col_field_order()) {
              if (fi == bf) {
                base_axis = BaseAxis::Col;
                break;
              }
            }
          }
        } else {
          // No base field set: fall back to the row axis if it is
          // single-level and has more than one leaf, else the col
          // axis if single-level, else give up.
          if (table.row_field_order().size() == 1 && actual_row_count > 1) {
            base_axis = BaseAxis::Row;
          } else if (table.col_field_order().size() == 1 && actual_col_count > 1) {
            base_axis = BaseAxis::Col;
          }
        }
        // MVP scope: only single-level base axis supported.
        if (base_axis == BaseAxis::Row && table.row_field_order().size() != 1) {
          base_axis = BaseAxis::None;
        }
        if (base_axis == BaseAxis::Col && table.col_field_order().size() != 1) {
          base_axis = BaseAxis::None;
        }
        if (base_axis == BaseAxis::None) {
          break;
        }
        const std::size_t axis_n = base_axis == BaseAxis::Row ? actual_row_count : actual_col_count;
        // Build base_pos[p] -> optional reference position along the
        // base axis. Sentinels resolve to (p-1) / (p+1); a specific
        // item index resolves to the leaf whose label matches the
        // base field's `items[index].name`.
        std::vector<std::optional<std::size_t>> base_pos(axis_n);
        const std::uint32_t base_item = df.show_as_base_item.value_or(kShowAsBasePrev);
        if (base_item == kShowAsBasePrev) {
          for (std::size_t p = 1; p < axis_n; ++p) {
            base_pos[p] = p - 1;
          }
        } else if (base_item == kShowAsBaseNext) {
          for (std::size_t p = 0; p + 1 < axis_n; ++p) {
            base_pos[p] = p + 1;
          }
        } else {
          // Specific item: `baseItem` is a cache shared-items index, the
          // same space as `<item x="N">`. Resolve it to a base-field item
          // by matching `cache_index` rather than the item's position in
          // `items` — `items` excludes subtotal / grand-total markers, so a
          // positional lookup shifts by the number of preceding markers.
          std::optional<std::size_t> fixed;
          if (df.show_as_base_field.has_value()) {
            const std::uint32_t bf = *df.show_as_base_field;
            if (bf < table.fields().size()) {
              const auto& items = table.fields()[bf].items;
              const std::string* target = nullptr;
              for (std::size_t j = 0; j < items.size(); ++j) {
                const bool hit = items[j].has_cache_index ? (items[j].cache_index == base_item) : (j == base_item);
                if (hit) {
                  target = &items[j].name;
                  break;
                }
              }
              if (target != nullptr) {
                if (base_axis == BaseAxis::Row) {
                  for (std::size_t p = 0; p < result.rows.size() && p < axis_n; ++p) {
                    if (result.rows[p].label == *target) {
                      fixed = p;
                      break;
                    }
                  }
                } else {
                  for (std::size_t p = 0; p < result.cols.size() && p < axis_n; ++p) {
                    if (result.cols[p].label == *target) {
                      fixed = p;
                      break;
                    }
                  }
                }
              }
            }
          }
          if (fixed.has_value()) {
            for (std::size_t p = 0; p < axis_n; ++p) {
              base_pos[p] = *fixed;
            }
          }
          // No match -> all base_pos remain nullopt; every cell
          // becomes blank, matching Excel's behaviour for an
          // unresolved base item.
        }
        const bool percent = (mode == ShowValuesAs::PercentDifferenceFrom);
        // Snapshot original numeric cells before mutating, so a
        // mutated cell never serves as somebody else's base reference.
        // Layout: [r][c] -> {has_value, number}.
        std::vector<std::vector<std::pair<bool, double>>> snapshot(
            actual_row_count, std::vector<std::pair<bool, double>>(actual_col_count, {false, 0.0}));
        for (std::size_t r = 0; r < actual_row_count; ++r) {
          for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
            if (df_idx >= result.values[r][c].size()) {
              continue;
            }
            snapshot[r][c] = cell_num(result.values[r][c][df_idx]);
          }
        }
        for (std::size_t r = 0; r < actual_row_count; ++r) {
          for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
            if (df_idx >= result.values[r][c].size()) {
              continue;
            }
            const std::size_t p = base_axis == BaseAxis::Row ? r : c;
            Value& cell = result.values[r][c][df_idx];
            const auto& cur = snapshot[r][c];
            if (!cur.first) {
              continue;
            }
            if (!base_pos[p].has_value()) {
              cell = Value::blank();
              continue;
            }
            const std::size_t bp = *base_pos[p];
            const std::size_t br = base_axis == BaseAxis::Row ? bp : r;
            const std::size_t bc = base_axis == BaseAxis::Col ? bp : c;
            if (br >= snapshot.size() || bc >= snapshot[br].size()) {
              cell = Value::blank();
              continue;
            }
            const auto& base = snapshot[br][bc];
            if (!base.first) {
              cell = Value::blank();
              continue;
            }
            if (percent) {
              if (base.second == 0.0) {
                cell = Value::error(ErrorCode::Div0);
              } else {
                cell = Value::number(cur.second / base.second - 1.0);
              }
            } else {
              cell = Value::number(cur.second - base.second);
            }
          }
        }
        break;
      }
      case ShowValuesAs::PercentOfParentRow:
      case ShowValuesAs::PercentOfParentCol:
      case ShowValuesAs::PercentOfParent: {
        // Resolve which axis hosts the parent and which depth (in
        // that axis's field-order) the parent field sits at. For
        // PercentOfParent the axis is determined by which order the
        // base field belongs to; for the *Row / *Col variants the
        // axis is fixed and `base_field` is optional (defaults to
        // immediate parent for multi-level, grand total otherwise).
        const PivotDataField& df = table.data_fields()[df_idx];
        enum class Axis { None, Row, Col } parent_axis = Axis::None;
        std::optional<std::size_t> base_depth;
        auto find_in = [&](const std::vector<std::uint32_t>& order, std::uint32_t fi) -> std::optional<std::size_t> {
          for (std::size_t i = 0; i < order.size(); ++i) {
            if (order[i] == fi) {
              return i;
            }
          }
          return std::nullopt;
        };
        if (mode == ShowValuesAs::PercentOfParent) {
          if (df.show_as_base_field.has_value()) {
            auto rd = find_in(table.row_field_order(), *df.show_as_base_field);
            if (rd.has_value()) {
              parent_axis = Axis::Row;
              base_depth = rd;
            } else {
              auto cd = find_in(table.col_field_order(), *df.show_as_base_field);
              if (cd.has_value()) {
                parent_axis = Axis::Col;
                base_depth = cd;
              }
            }
          }
        } else if (mode == ShowValuesAs::PercentOfParentRow) {
          parent_axis = Axis::Row;
          if (df.show_as_base_field.has_value()) {
            base_depth = find_in(table.row_field_order(), *df.show_as_base_field);
          }
        } else {  // PercentOfParentCol
          parent_axis = Axis::Col;
          if (df.show_as_base_field.has_value()) {
            base_depth = find_in(table.col_field_order(), *df.show_as_base_field);
          }
        }
        if (parent_axis == Axis::None) {
          break;
        }
        // Build per-leaf parent total along the chosen axis.
        // Strategy: for each leaf p, find the row_subtotal/col_subtotal
        // whose `depth == base_depth` AND whose leaf-set contains p.
        // If `base_depth` is unset, fall back to the deepest enclosing
        // subtotal (for a single-level axis there is none → use the
        // grand total).
        const auto& subs_leaf_sets = parent_axis == Axis::Row ? row_subtotal_leaf_sets : col_subtotal_leaf_sets;
        const std::size_t axis_n = parent_axis == Axis::Row ? actual_row_count : actual_col_count;
        std::vector<std::optional<double>> parent_total(axis_n);
        // Compute the grand total as a fallback denominator.
        std::optional<double> grand;
        if (df_idx < result.grand_totals.size()) {
          auto [ok, n] = cell_num(result.grand_totals[df_idx]);
          if (ok) {
            grand = n;
          }
        }
        if (!grand.has_value()) {
          double t = 0.0;
          bool any = false;
          for (std::size_t r = 0; r < actual_row_count; ++r) {
            for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
              if (df_idx >= result.values[r][c].size()) {
                continue;
              }
              auto [ok, n] = cell_num(result.values[r][c][df_idx]);
              if (ok) {
                t += n;
                any = true;
              }
            }
          }
          if (any) {
            grand = t;
          }
        }
        // Pick the right subtotal-values vector for the parent axis.
        for (std::size_t p = 0; p < axis_n; ++p) {
          std::optional<std::size_t> chosen_sub;
          std::size_t chosen_depth = 0;
          for (std::size_t s = 0; s < subs_leaf_sets.size(); ++s) {
            const auto& set = subs_leaf_sets[s];
            if (std::find(set.begin(), set.end(), p) == set.end()) {
              continue;
            }
            std::uint32_t sub_depth = 0;
            if (parent_axis == Axis::Row && s < result.row_subtotals.size()) {
              sub_depth = result.row_subtotals[s].depth;
            } else if (parent_axis == Axis::Col && s < result.col_subtotals.size()) {
              sub_depth = result.col_subtotals[s].depth;
            }
            if (base_depth.has_value()) {
              if (sub_depth == *base_depth) {
                chosen_sub = s;
                break;
              }
            } else {
              if (!chosen_sub.has_value() || sub_depth >= chosen_depth) {
                chosen_sub = s;
                chosen_depth = sub_depth;
              }
            }
          }
          if (chosen_sub.has_value()) {
            if (parent_axis == Axis::Row) {
              const RowSubtotal& sub = result.row_subtotals[*chosen_sub];
              if (df_idx < sub.values.size()) {
                auto [ok, n] = cell_num(sub.values[df_idx]);
                if (ok) {
                  parent_total[p] = n;
                }
              }
            } else {
              // Col subtotal "row total" for a leaf row is the sum
              // across that leaf's row in the col_subtotal columns.
              // But here we need the col-axis parent total at leaf
              // col `p`: it is the sum across the col_subtotal whose
              // leaf set contains p, taken over all row leaves. We
              // surface that as the sum of the col_subtotal's per-row
              // values at this df.
              const ColSubtotal& sub = result.col_subtotals[*chosen_sub];
              double t = 0.0;
              bool any = false;
              for (const auto& row_slot : sub.values) {
                if (df_idx >= row_slot.size()) {
                  continue;
                }
                auto [ok, n] = cell_num(row_slot[df_idx]);
                if (ok) {
                  t += n;
                  any = true;
                }
              }
              if (any) {
                parent_total[p] = t;
              }
            }
          } else if (grand.has_value()) {
            parent_total[p] = grand;
          }
        }
        // Apply the transform. Only the leaf cells are scaled; the
        // subtotal / grand-total cells stay at their raw aggregate
        // (consistent with how RunningTotal / Index leave them alone).
        for (std::size_t r = 0; r < actual_row_count; ++r) {
          for (std::size_t c = 0; c < actual_col_count && c < result.values[r].size(); ++c) {
            if (df_idx >= result.values[r][c].size()) {
              continue;
            }
            const std::size_t p = parent_axis == Axis::Row ? r : c;
            if (!parent_total[p].has_value()) {
              continue;
            }
            scale_cell(result.values[r][c][df_idx], *parent_total[p]);
          }
        }
        break;
      }
      case ShowValuesAs::Normal:
        break;
    }
  }

  // Re-sync the legacy mirrors after the transform pass: the flat
  // `result.subtotals[i][df]` view tracks `result.row_subtotals[i].values[df]`,
  // and `result.grand_total` tracks `result.grand_totals[0]`.
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

}  // namespace formulon::pivot
