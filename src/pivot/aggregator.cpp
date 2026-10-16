
#include "pivot/aggregator.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <vector>

#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "utils/cancellation_snap.h"
#include "utils/error.h"
#include "utils/subnormal_flush.h"
#include "value.h"

namespace formulon::pivot {

std::size_t RecordIndexView::record_at(std::size_t position) const {
  if (spans_ == nullptr) {
    return direct_data_ != nullptr && position < direct_count_ ? direct_data_[position] : 0U;
  }
  if (position >= total_count_ || span_count_ == 0U) {
    return 0U;
  }

  std::size_t span_index = 0U;
  if (position >= cached_begin_ && position < cached_end_ && cached_span_ < span_count_) {
    span_index = cached_span_;
  } else if (position >= cached_end_ && cached_span_ < span_count_) {
    // Source-order reads advance to the next span in amortized O(1).
    span_index = cached_span_;
    while (span_index < span_count_ && position >= spans_[span_index].cumulative_end) {
      ++span_index;
    }
  } else {
    // A restart or random read binary-searches the cumulative ends.
    std::size_t begin = 0U;
    std::size_t end = span_count_;
    while (begin < end) {
      const std::size_t middle = begin + (end - begin) / 2U;
      if (position < spans_[middle].cumulative_end) {
        end = middle;
      } else {
        begin = middle + 1U;
      }
    }
    span_index = begin;
  }
  if (span_index >= span_count_) {
    return 0U;
  }
  const RecordIndexSpan& span = spans_[span_index];
  cached_span_ = span_index;
  cached_end_ = span.cumulative_end;
  cached_begin_ = cached_end_ - span.count;
  return span.data[position - cached_begin_];
}

namespace {

// A non-finite aggregate is #NUM!, as Excel reports it.
Value finite_number(double value);

/// Accumulator behind every non-Sum aggregation. A leaf adds its records in
/// source order with Excel's per-step snap; a total across both axes combines
/// the per-row states raw (no snap, no flush).
class AggregateState {
 public:
  explicit AggregateState(Aggregation aggregation) : aggregation_(aggregation) {}

  void add(const Value& value) {
    if (value.is_blank()) {
      return;
    }
    ++non_blank_;
    if (value.is_error()) {
      if (!error_.is_error()) {
        error_ = value;
      }
      return;
    }
    if (!value.is_number()) {
      return;
    }
    const double x = value.as_number();
    if (numbers_ == 0U) {
      s_ = x;
    } else {
      switch (aggregation_) {
        case Aggregation::Product:
          s_ = snap_cancellation(x, flush_subnormal(s_ * x));
          break;
        case Aggregation::Min:
          s_ = snap_cancellation(x, std::min(s_, x));
          break;
        case Aggregation::Max:
          s_ = snap_cancellation(x, std::max(s_, x));
          break;
        default:
          s_ = snap_cancellation(x, s_ + x);
          break;
      }
    }
    s2_ += x * x;
    ++numbers_;
  }

  void combine(const AggregateState& other) {
    non_blank_ += other.non_blank_;
    if (!error_.is_error() && other.error_.is_error()) {
      error_ = other.error_;
    }
    s2_ += other.s2_;
    if (other.numbers_ > 0U) {
      if (numbers_ == 0U) {
        s_ = other.s_;
      } else {
        switch (aggregation_) {
          case Aggregation::Product:
            s_ = s_ * other.s_;
            break;
          case Aggregation::Min:
            s_ = std::min(s_, other.s_);
            break;
          case Aggregation::Max:
            s_ = std::max(s_, other.s_);
            break;
          default:
            s_ = s_ + other.s_;
            break;
        }
      }
      numbers_ += other.numbers_;
    }
  }

  /// Zeroes the statistics a 2-D combination did not carry.
  void drop(const MissingStats& missing) {
    if (missing.count) {
      numbers_ = 0U;
    }
    if (missing.sum_of_squares) {
      s2_ = 0.0;
    }
  }

  /// True when no non-blank record contributed: the cell renders blank.
  bool all_blank() const { return non_blank_ == 0U; }

  Value finish() const {
    if (aggregation_ == Aggregation::Count) {
      return Value::number(static_cast<double>(non_blank_));
    }
    if (aggregation_ == Aggregation::CountNumbers) {
      return Value::number(static_cast<double>(numbers_));
    }
    if (error_.is_error()) {
      return error_;
    }
    const double n = static_cast<double>(numbers_);
    switch (aggregation_) {
      case Aggregation::Average:
        return numbers_ == 0U ? Value::error(ErrorCode::Div0) : finite_number(s_ / n);
      case Aggregation::Sum:
      case Aggregation::Product:
      case Aggregation::Min:
      case Aggregation::Max:
        return finite_number(numbers_ > 0U ? s_ : 0.0);
      case Aggregation::Var:
      case Aggregation::StdDev: {
        if (numbers_ < 2U) {
          return Value::error(ErrorCode::Div0);
        }
        return dispersion((s2_ - s_ * s_ / n) / (n - 1.0), aggregation_ == Aggregation::StdDev);
      }
      case Aggregation::VarP:
      case Aggregation::StdDevP: {
        if (numbers_ == 0U) {
          return Value::error(ErrorCode::Div0);
        }
        return dispersion((n * s2_ - s_ * s_) / (n * n), aggregation_ == Aggregation::StdDevP);
      }
      default:
        return Value::error(ErrorCode::NA);
    }
  }

 private:
  // Overflow (NaN or infinity) is #NUM!; a negative one-pass residue is 0.
  static Value dispersion(double variance, bool root) {
    if (!std::isfinite(variance)) {
      return Value::error(ErrorCode::Num);
    }
    variance = std::max(0.0, variance);
    return Value::number(root ? std::sqrt(variance) : variance);
  }

  Aggregation aggregation_;
  std::size_t non_blank_ = 0;
  std::size_t numbers_ = 0;
  Value error_ = Value::blank();
  double s_ = 0.0;
  double s2_ = 0.0;
};

/// Borrowed source of Values. The getter may resolve a cache record or read
/// an already-materialized vector; aggregation never stores the returned
/// Value beyond the current synchronous pass.
struct ValueInputView {
  const void* context = nullptr;
  std::size_t size = 0;
  Value (*get)(const void* context, std::size_t index) = nullptr;

  Value at(std::size_t index) const { return get == nullptr ? Value::blank() : get(context, index); }
};

Value vector_value_get(const void* context, std::size_t index) {
  const auto& values = *static_cast<const std::vector<Value>*>(context);
  return values[index];
}

struct RecordValueContext {
  const PivotCache* cache = nullptr;
  const RecordIndexView* records = nullptr;
  std::uint32_t field_index = 0;
};

Value record_cell(const PivotCache& cache, std::size_t record_index, std::uint32_t field_index) {
  if (record_index >= cache.records().size()) {
    return Value::blank();
  }
  return cell_value(cache, cache.records()[record_index], field_index);
}

Value record_value_get(const void* context, std::size_t index) {
  const auto& source = *static_cast<const RecordValueContext*>(context);
  return record_cell(*source.cache, source.records->record_at(index), source.field_index);
}

ValueInputView vector_value_view(const std::vector<Value>& values) {
  return ValueInputView{&values, values.size(), &vector_value_get};
}

ValueInputView record_value_view(const PivotCache& cache, const RecordIndexView& records, std::uint32_t field_index,
                                 RecordValueContext& context) {
  context = RecordValueContext{&cache, &records, field_index};
  return ValueInputView{&context, records.size(), &record_value_get};
}

// A non-finite aggregate is #NUM!, as Excel reports it.
Value finite_number(double value) {
  if (!std::isfinite(value)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(value);
}

Value aggregate_sum(const ValueInputView& values) {
  double sum = 0.0;
  for (std::size_t index = 0; index < values.size; ++index) {
    const Value value = values.at(index);
    if (value.is_error()) {
      return value;
    }
    if (const std::optional<double> number = numeric_aggregate_value(value); number.has_value()) {
      sum = snapped_add(*number, sum);
    }
  }
  return finite_number(sum);
}

Value apply_value_aggregation(Aggregation aggregation, const ValueInputView& values) {
  if (aggregation == Aggregation::Sum) {
    return aggregate_sum(values);
  }
  AggregateState state(aggregation);
  for (std::size_t index = 0; index < values.size; ++index) {
    state.add(values.at(index));
  }
  return state.finish();
}

// ---------------------------------------------------------------------------
// Result-side text reification.
// ---------------------------------------------------------------------------
//
// `PivotResult::values` / `subtotals` / grand totals must outlive the
// cache they were computed against (GETPIVOTDATA reads them outside of
// any specific evaluation arena). Numbers, bools, errors, and blanks
// are trivially copyable. Text is the only kind that needs storage —
// we copy the bytes into `result.text_storage` and rebuild a `Value`
// pointing into the deque entry. Pointer/iterator stability of
// `std::deque` keeps the views valid across subsequent appends.
Value reify(const Value& value, PivotResult& result) {
  if (!value.is_text()) {
    return value;
  }
  result.text_storage.emplace_back(value.as_text());
  return Value::text(result.text_storage.back());
}

}  // namespace

Value apply_aggregation(Aggregation aggregation, const std::vector<Value>& values) {
  return apply_value_aggregation(aggregation, vector_value_view(values));
}

Value apply_record_aggregation(Aggregation aggregation, const PivotCache& cache, const RecordIndexView& records,
                               std::uint32_t field_index) {
  RecordValueContext context;
  const ValueInputView values = record_value_view(cache, records, field_index, context);
  return apply_value_aggregation(aggregation, values);
}

Value apply_record_aggregation(Aggregation aggregation, const PivotCache& cache,
                               const std::vector<std::size_t>& records, std::uint32_t field_index) {
  return apply_record_aggregation(aggregation, cache, RecordIndexView(records), field_index);
}

std::optional<double> numeric_aggregate_value(const Value& value) {
  if (value.is_number()) {
    return value.as_number();
  }
  return std::nullopt;
}

namespace {

/// Appends the bucket at (`row_leaf`, `col_leaf`) to `spans` when it exists and is non-empty.
void append_bucket_span(const RecordBuckets& buckets, std::size_t row_leaf, std::size_t col_leaf,
                        std::vector<RecordIndexSpan>& spans, std::size_t& total_count) {
  if (row_leaf >= buckets.size() || col_leaf >= buckets[row_leaf].size()) {
    return;
  }
  const std::vector<std::size_t>& records = buckets[row_leaf][col_leaf];
  if (records.empty()) {
    return;
  }
  total_count += records.size();
  spans.push_back(RecordIndexSpan{records.data(), records.size(), total_count});
}

}  // namespace

RecordIndexView make_leaf_set_record_view(const RecordBuckets& buckets, const std::vector<std::size_t>& row_leaves,
                                          const std::vector<std::size_t>& col_leaves,
                                          std::vector<RecordIndexSpan>& spans) {
  spans.clear();
  std::size_t total_count = 0U;
  for (const std::size_t col_leaf : col_leaves) {
    for (const std::size_t row_leaf : row_leaves) {
      append_bucket_span(buckets, row_leaf, col_leaf, spans, total_count);
    }
  }
  return RecordIndexView::from_spans(spans.data(), spans.size(), total_count);
}

RecordIndexView make_leaf_set_record_view(const RecordBuckets& buckets, std::size_t row_leaf,
                                          const std::vector<std::size_t>& col_leaves,
                                          std::vector<RecordIndexSpan>& spans) {
  spans.clear();
  std::size_t total_count = 0U;
  for (const std::size_t col_leaf : col_leaves) {
    append_bucket_span(buckets, row_leaf, col_leaf, spans, total_count);
  }
  return RecordIndexView::from_spans(spans.data(), spans.size(), total_count);
}

RecordIndexView make_leaf_set_record_view(const RecordBuckets& buckets, const std::vector<std::size_t>& row_leaves,
                                          std::size_t col_leaf, std::vector<RecordIndexSpan>& spans) {
  spans.clear();
  std::size_t total_count = 0U;
  for (const std::size_t row_leaf : row_leaves) {
    append_bucket_span(buckets, row_leaf, col_leaf, spans, total_count);
  }
  return RecordIndexView::from_spans(spans.data(), spans.size(), total_count);
}

Value aggregate_records_or_blank(Aggregation aggregation, const PivotCache& cache, const RecordIndexView& records,
                                 std::uint32_t field_index, PivotResult& result) {
  if (records.empty()) {
    return Value::blank();
  }
  AggregateState state(aggregation);
  for (std::size_t i = 0; i < records.size(); ++i) {
    state.add(record_cell(cache, records.record_at(i), field_index));
  }
  if (state.all_blank()) {
    return Value::blank();
  }
  // Sum keeps its own snapped chain unless its totals follow the accumulator model.
  return aggregation == Aggregation::Sum && !result.sum_totals_as_state
             ? reify(apply_record_aggregation(aggregation, cache, records, field_index), result)
             : state.finish();
}

Value aggregate_records_or_blank(Aggregation aggregation, const PivotCache& cache,
                                 const std::vector<std::size_t>& records, std::uint32_t field_index,
                                 PivotResult& result) {
  return aggregate_records_or_blank(aggregation, cache, RecordIndexView(records), field_index, result);
}

namespace {

bool keeps_count(Aggregation a) {
  return a == Aggregation::Average || a == Aggregation::CountNumbers || a == Aggregation::StdDev ||
         a == Aggregation::StdDevP || a == Aggregation::Var || a == Aggregation::VarP;
}

bool keeps_sum_of_squares(Aggregation a) {
  return a == Aggregation::StdDev || a == Aggregation::StdDevP || a == Aggregation::Var || a == Aggregation::VarP;
}

}  // namespace

MissingStats missing_custom_subtotal_stats(const PivotTable& table, std::size_t depth, Aggregation custom,
                                           Aggregation data_aggregation) {
  if (depth == 0U) {
    return {};
  }
  bool count = keeps_count(data_aggregation);
  bool squares = keeps_sum_of_squares(data_aggregation);
  for (std::size_t k = 0; k < depth && k < table.row_field_order().size(); ++k) {
    const std::uint32_t field_index = table.row_field_order()[k];
    if (field_index >= table.fields().size()) {
      continue;
    }
    for (const SubtotalFn fn : table.fields()[field_index].subtotal_fns) {
      const Aggregation above = static_cast<Aggregation>(static_cast<std::uint8_t>(fn));
      count = count || keeps_count(above);
      squares = squares || keeps_sum_of_squares(above);
    }
  }
  return {keeps_count(custom) && !count, keeps_sum_of_squares(custom) && !squares};
}

bool sum_totals_follow_accumulator(const PivotTable& table) {
  for (const std::uint32_t field_index : table.row_field_order()) {
    if (field_index >= table.fields().size()) {
      continue;
    }
    for (const SubtotalFn fn : table.fields()[field_index].subtotal_fns) {
      if (fn != SubtotalFn::Sum) {
        return true;
      }
    }
  }
  return false;
}

Value aggregate_row_partials_or_blank(Aggregation aggregation, const PivotCache& cache, const RecordBuckets& buckets,
                                      const std::vector<std::size_t>& row_leaves,
                                      const std::vector<std::size_t>& col_leaves, std::uint32_t field_index,
                                      PivotResult& result, const MissingStats& missing) {
  if (aggregation == Aggregation::Sum && !result.sum_totals_as_state) {
    std::vector<RecordIndexSpan> spans;
    return aggregate_records_or_blank(
        aggregation, cache, make_leaf_set_record_view(buckets, row_leaves, col_leaves, spans), field_index, result);
  }
  AggregateState total(aggregation);
  for (const std::size_t row_leaf : row_leaves) {
    AggregateState row(aggregation);
    for (const std::size_t col_leaf : col_leaves) {
      if (row_leaf >= buckets.size() || col_leaf >= buckets[row_leaf].size()) {
        continue;
      }
      for (const std::size_t record : buckets[row_leaf][col_leaf]) {
        row.add(record_cell(cache, record, field_index));
      }
    }
    total.combine(row);
  }
  // One column leaf is a 1-D total and exact.
  if (col_leaves.size() > 1U) {
    total.drop(missing);
  }
  return total.all_blank() ? Value::blank() : total.finish();
}

namespace {

/// Running total of rendered cells: cancellation-snapped, or a plain IEEE sum.
class CellSum {
 public:
  explicit CellSum(bool snap) : snap_(snap) {}

  void add(const Value& cell) {
    if (cell.is_blank()) {
      return;
    }
    if (!cell.is_number()) {
      non_numeric_ = true;
      return;
    }
    const double x = cell.as_number();
    total_ = any_ ? (snap_ ? snapped_add(x, total_) : total_ + x) : x;
    any_ = true;
  }

  /// Leaves `slot` alone when no cell contributed or a cell was not a number.
  void store(Value& slot) const {
    if (any_ && !non_numeric_) {
      slot = Value::number(total_);
    }
  }

 private:
  bool snap_;
  bool any_ = false;
  bool non_numeric_ = false;
  double total_ = 0.0;
};

}  // namespace

void recompute_sum_totals_from_cells(const PivotTable& table, PivotResult& result,
                                     const std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                                     const std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets) {
  const std::size_t row_count = result.values.size();
  const std::size_t col_count = row_count == 0U ? 0U : result.values[0].size();
  const bool has_rows = !table.row_field_order().empty();
  const bool has_cols = !table.col_field_order().empty();
  const std::size_t col_subtotal_count = std::min(result.col_subtotals.size(), col_subtotal_leaf_sets.size());
  const std::size_t row_subtotal_count = std::min(result.row_subtotals.size(), row_subtotal_leaf_sets.size());

  for (std::size_t df = 0; df < table.data_fields().size(); ++df) {
    if (table.data_fields()[df].aggregation != Aggregation::Sum || result.sum_totals_as_state) {
      continue;
    }
    auto is_sum = [](const std::optional<Aggregation>& spec) {
      return spec.value_or(Aggregation::Sum) == Aggregation::Sum;
    };
    auto cell = [&](std::size_t r, std::size_t c) -> const Value& { return result.values[r][c][df]; };

    if (result.row_leaf_totals.size() == row_count) {
      for (std::size_t r = 0; r < row_count; ++r) {
        CellSum sum(/*snap=*/true);
        for (std::size_t c = 0; c < col_count; ++c) {
          sum.add(cell(r, c));
        }
        sum.store(result.row_leaf_totals[r][df]);
      }
    }
    if (result.col_leaf_totals.size() == col_count) {
      for (std::size_t c = 0; c < col_count; ++c) {
        CellSum sum(/*snap=*/true);
        for (std::size_t r = 0; r < row_count; ++r) {
          sum.add(cell(r, c));
        }
        sum.store(result.col_leaf_totals[c][df]);
      }
    }

    for (std::size_t cs = 0; cs < col_subtotal_count; ++cs) {
      ColSubtotal& col_subtotal = result.col_subtotals[cs];
      if (!is_sum(col_subtotal.aggregation) || col_subtotal.values.size() != row_count) {
        continue;
      }
      CellSum total(/*snap=*/false);
      for (std::size_t r = 0; r < row_count; ++r) {
        CellSum sum(/*snap=*/true);
        for (const std::size_t c : col_subtotal_leaf_sets[cs]) {
          sum.add(cell(r, c));
        }
        sum.store(col_subtotal.values[r][df]);
        total.add(col_subtotal.values[r][df]);
      }
      total.store(col_subtotal.total[df]);
    }

    for (std::size_t rs = 0; rs < row_subtotal_count; ++rs) {
      RowSubtotal& row_subtotal = result.row_subtotals[rs];
      if (!is_sum(row_subtotal.aggregation)) {
        continue;
      }
      const std::vector<std::size_t>& rows = row_subtotal_leaf_sets[rs];
      for (std::size_t c = 0; c < row_subtotal.col_values.size() && c < col_count; ++c) {
        CellSum sum(/*snap=*/true);
        for (const std::size_t r : rows) {
          sum.add(cell(r, c));
        }
        sum.store(row_subtotal.col_values[c][df]);
      }
      // Without a column field the subtotal sits on a single leaf column.
      CellSum sum(/*snap=*/!has_cols);
      for (const std::size_t r : rows) {
        sum.add(has_cols ? result.row_leaf_totals[r][df] : cell(r, 0));
      }
      sum.store(row_subtotal.values[df]);
      if (rs < result.subtotals.size()) {
        result.subtotals[rs][df] = row_subtotal.values[df];
      }
      for (std::size_t cs = 0; cs < row_subtotal.col_subtotal_values.size() && cs < col_subtotal_count; ++cs) {
        const ColSubtotal& col_subtotal = result.col_subtotals[cs];
        if (!is_sum(col_subtotal.aggregation)) {
          continue;
        }
        CellSum both(/*snap=*/false);
        for (const std::size_t r : rows) {
          both.add(col_subtotal.values[r][df]);
        }
        both.store(row_subtotal.col_subtotal_values[cs][df]);
      }
    }

    if (df < result.grand_totals.size()) {
      CellSum total(/*snap=*/!(has_rows && has_cols));
      if (has_rows && has_cols) {
        for (std::size_t r = 0; r < row_count; ++r) {
          total.add(result.row_leaf_totals[r][df]);
        }
      } else {
        for (std::size_t r = 0; r < row_count; ++r) {
          for (std::size_t c = 0; c < col_count; ++c) {
            total.add(cell(r, c));
          }
        }
      }
      total.store(result.grand_totals[df]);
    }
  }
  if (!result.grand_totals.empty()) {
    result.grand_total = result.grand_totals[0];
  }
}

}  // namespace formulon::pivot
