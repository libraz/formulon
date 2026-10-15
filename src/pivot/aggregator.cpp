
#include "pivot/aggregator.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <vector>

#include "numeric_aggregate_kernels.h"
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

Value record_value_get(const void* context, std::size_t index) {
  const auto& source = *static_cast<const RecordValueContext*>(context);
  const std::size_t record_index = source.records->record_at(index);
  if (record_index >= source.cache->records().size()) {
    return Value::blank();
  }
  return cell_value(*source.cache, source.cache->records()[record_index], source.field_index);
}

ValueInputView vector_value_view(const std::vector<Value>& values) {
  return ValueInputView{&values, values.size(), &vector_value_get};
}

ValueInputView record_value_view(const PivotCache& cache, const RecordIndexView& records, std::uint32_t field_index,
                                 RecordValueContext& context) {
  context = RecordValueContext{&cache, &records, field_index};
  return ValueInputView{&context, records.size(), &record_value_get};
}

Value lift_kernel_result(Expected<double, ErrorCode> result) {
  if (!result) {
    return Value::error(result.error());
  }
  return Value::number(result.value());
}

// A non-finite aggregate is #NUM!, as Excel reports it.
Value finite_number(double value) {
  if (!std::isfinite(value)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(value);
}

// Variance and standard deviation use the contiguous NumericKernel so their
// two-pass arithmetic keeps the established rounding order. The input scan
// also checks errors, avoiding a separate callback pass before materializing
// the numeric slice.
struct NumericValues {
  std::vector<double> values;
  std::optional<Value> error;
};

NumericValues collect_numeric_values(const ValueInputView& input) {
  NumericValues result;
  result.values.reserve(input.size);
  for (std::size_t index = 0; index < input.size; ++index) {
    const Value value = input.at(index);
    if (value.is_error()) {
      result.error = value;
      break;
    }
    if (const std::optional<double> number = numeric_aggregate_value(value); number.has_value()) {
      result.values.push_back(*number);
    }
  }
  return result;
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

// Excel's pivot `Count` mirrors COUNTA: any non-blank cell counts,
// including text, booleans and errors.
Value aggregate_count(const ValueInputView& values) {
  double count = 0.0;
  for (std::size_t index = 0; index < values.size; ++index) {
    if (!values.at(index).is_blank()) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

// Excel's pivot `CountNumbers` mirrors COUNT: only numeric cells
// (booleans included, per Excel). Errors are classified, never propagated.
Value aggregate_count_numbers(const ValueInputView& values) {
  double count = 0.0;
  for (std::size_t index = 0; index < values.size; ++index) {
    const Value value = values.at(index);
    if (value.is_number() || value.is_boolean()) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

Value aggregate_average(const ValueInputView& values) {
  double sum = 0.0;
  std::size_t count = 0U;
  for (std::size_t index = 0; index < values.size; ++index) {
    const Value value = values.at(index);
    if (value.is_error()) {
      return value;
    }
    if (const std::optional<double> number = numeric_aggregate_value(value); number.has_value()) {
      sum = snapped_add(*number, sum);
      ++count;
    }
  }
  if (count == 0U) {
    return Value::error(ErrorCode::Div0);
  }
  return finite_number(sum / static_cast<double>(count));
}

Value aggregate_max(const ValueInputView& values) {
  double maximum = -std::numeric_limits<double>::infinity();
  std::size_t count = 0U;
  for (std::size_t index = 0; index < values.size; ++index) {
    const Value value = values.at(index);
    if (value.is_error()) {
      return value;
    }
    if (const std::optional<double> number = numeric_aggregate_value(value); number.has_value()) {
      const double x = *number;
      maximum = (count == 0U || x > maximum) ? x : maximum;
      ++count;
    }
  }
  // Excel's pivot MAX over an empty/all-text group returns 0.
  return finite_number(count > 0U ? maximum : 0.0);
}

Value aggregate_min(const ValueInputView& values) {
  double minimum = std::numeric_limits<double>::infinity();
  std::size_t count = 0U;
  for (std::size_t index = 0; index < values.size; ++index) {
    const Value value = values.at(index);
    if (value.is_error()) {
      return value;
    }
    if (const std::optional<double> number = numeric_aggregate_value(value); number.has_value()) {
      const double x = *number;
      minimum = (count == 0U || x < minimum) ? x : minimum;
      ++count;
    }
  }
  return finite_number(count > 0U ? minimum : 0.0);
}

Value variance_helper(const ValueInputView& values, bool population) {
  NumericValues numeric = collect_numeric_values(values);
  if (numeric.error.has_value()) {
    return *numeric.error;
  }
  // Booleans count as 0/1, text and blanks are skipped; too few values yield
  // #DIV/0! as in Excel's pivot.
  return lift_kernel_result(numeric_aggregate_kernels::run_variance(numeric.values, !population));
}

Value aggregate_stddev(const ValueInputView& values, bool population) {
  NumericValues numeric = collect_numeric_values(values);
  if (numeric.error.has_value()) {
    return *numeric.error;
  }
  return lift_kernel_result(numeric_aggregate_kernels::run_stdev(numeric.values, !population));
}

Value aggregate_product(const ValueInputView& values) {
  double product = 1.0;
  std::size_t count = 0U;
  for (std::size_t index = 0; index < values.size; ++index) {
    const Value value = values.at(index);
    if (value.is_error()) {
      return value;
    }
    if (const std::optional<double> number = numeric_aggregate_value(value); number.has_value()) {
      // Like PRODUCT, start from the first value and flush a subnormal step.
      product = count > 0U ? flush_subnormal(product * *number) : *number;
      ++count;
    }
  }
  // Excel's pivot PRODUCT on an empty/all-text group returns 0.
  return finite_number(count > 0U ? product : 0.0);
}

Value apply_value_aggregation(Aggregation aggregation, const ValueInputView& values) {
  switch (aggregation) {
    case Aggregation::Sum:
      return aggregate_sum(values);
    case Aggregation::Count:
      return aggregate_count(values);
    case Aggregation::Average:
      return aggregate_average(values);
    case Aggregation::Max:
      return aggregate_max(values);
    case Aggregation::Min:
      return aggregate_min(values);
    case Aggregation::Product:
      return aggregate_product(values);
    case Aggregation::CountNumbers:
      return aggregate_count_numbers(values);
    case Aggregation::StdDev:
      return aggregate_stddev(values, /*population=*/false);
    case Aggregation::StdDevP:
      return aggregate_stddev(values, /*population=*/true);
    case Aggregation::Var:
      return variance_helper(values, /*population=*/false);
    case Aggregation::VarP:
      return variance_helper(values, /*population=*/true);
  }
  return Value::error(ErrorCode::NA);
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
  if (value.is_boolean()) {
    return value.as_boolean() ? 1.0 : 0.0;
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
  return reify(apply_record_aggregation(aggregation, cache, records, field_index), result);
}

Value aggregate_records_or_blank(Aggregation aggregation, const PivotCache& cache,
                                 const std::vector<std::size_t>& records, std::uint32_t field_index,
                                 PivotResult& result) {
  return aggregate_records_or_blank(aggregation, cache, RecordIndexView(records), field_index, result);
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
    if (table.data_fields()[df].aggregation != Aggregation::Sum) {
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
