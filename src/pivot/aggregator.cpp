
#include "pivot/aggregator.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <vector>

#include "numeric_aggregate_kernels.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// Returns the first `Value::error` found in `values`, or `nullptr`.
const Value* first_error(const std::vector<Value>& values) {
  for (const auto& v : values) {
    if (v.is_error()) {
      return &v;
    }
  }
  return nullptr;
}

// Numeric coercion for arithmetic aggregations. Booleans coerce; text
// is skipped (Excel's SUM/MAX/MIN over a Value column ignore text).
// `out` receives the coerced number on success.
bool coerce_arithmetic(const Value& v, double& out) noexcept {
  switch (v.kind()) {
    case ValueKind::Number:
      out = v.as_number();
      return true;
    case ValueKind::Bool:
      out = v.as_boolean() ? 1.0 : 0.0;
      return true;
    default:
      return false;
  }
}

struct ArithmeticSummary {
  double sum = 0.0;
  double product = 1.0;
  double min = std::numeric_limits<double>::infinity();
  double max = -std::numeric_limits<double>::infinity();
  std::size_t count = 0;
};

ArithmeticSummary summarize_arithmetic(const std::vector<Value>& values) {
  ArithmeticSummary summary;
  for (const auto& v : values) {
    double x = 0.0;
    if (!coerce_arithmetic(v, x)) {
      continue;
    }
    summary.sum += x;
    summary.product *= x;
    summary.min = (summary.count == 0 || x < summary.min) ? x : summary.min;
    summary.max = (summary.count == 0 || x > summary.max) ? x : summary.max;
    ++summary.count;
  }
  return summary;
}

std::optional<double> numeric_view_get(const void* context, std::size_t index) {
  const auto& values = *static_cast<const std::vector<Value>*>(context);
  double number = 0.0;
  if (!coerce_arithmetic(values[index], number)) {
    return std::nullopt;
  }
  return number;
}

numeric_aggregate_kernels::NumericInputView numeric_view(const std::vector<Value>& values) {
  return numeric_aggregate_kernels::NumericInputView{&values, values.size(), &numeric_view_get};
}

Value lift_kernel_result(Expected<double, ErrorCode> result) {
  if (!result) {
    return Value::error(result.error());
  }
  return Value::number(result.value());
}

// A non-finite aggregate is #NUM!, as Excel reports it.
Value finite_number(double v) {
  if (std::isnan(v) || std::isinf(v)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(v);
}

Value AggregateSum(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  return finite_number(summarize_arithmetic(values).sum);
}

// Excel's pivot `Count` mirrors COUNTA: any non-blank cell counts,
// including text, booleans and errors.
//
// The COUNT family classifies cells rather than coercing their values, so
// an error in the group is a cell like any other and never becomes the
// result. A single `#N/A` left by a failed lookup in the value column
// would otherwise turn the group's count, its subtotal and the grand
// total into `#N/A`, where Excel reports the count.
Value AggregateCount(const std::vector<Value>& values) {
  double count = 0.0;
  for (const auto& v : values) {
    if (!v.is_blank()) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

// Excel's pivot `CountNumbers` mirrors COUNT: only numeric cells
// (booleans included, per Excel). Errors are cells that do not qualify,
// not a result -- the same classification rule `AggregateCount` follows.
Value AggregateCountNumbers(const std::vector<Value>& values) {
  double count = 0.0;
  for (const auto& v : values) {
    if (v.is_number() || v.is_boolean()) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

Value AggregateAverage(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  return lift_kernel_result(numeric_aggregate_kernels::run_average(numeric_view(values)));
}

Value AggregateMax(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  const ArithmeticSummary summary = summarize_arithmetic(values);
  // Excel's pivot MAX over an empty/all-text group returns 0.
  return finite_number(summary.count > 0 ? summary.max : 0.0);
}

Value AggregateMin(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  const ArithmeticSummary summary = summarize_arithmetic(values);
  return finite_number(summary.count > 0 ? summary.min : 0.0);
}

Value variance_helper(const std::vector<Value>& values, bool population) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  // Booleans count as 0/1, text and blanks are skipped; too few values yield #DIV/0! as in Excel's pivot.
  return lift_kernel_result(numeric_aggregate_kernels::run_variance(numeric_view(values), !population));
}

Value AggregateVar(const std::vector<Value>& values) {
  return variance_helper(values, /*population=*/false);
}

Value AggregateVarP(const std::vector<Value>& values) {
  return variance_helper(values, /*population=*/true);
}

Value stddev_helper(const std::vector<Value>& values, bool population) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  return lift_kernel_result(numeric_aggregate_kernels::run_stdev(numeric_view(values), !population));
}

Value AggregateStdDev(const std::vector<Value>& values) {
  return stddev_helper(values, /*population=*/false);
}

Value AggregateStdDevP(const std::vector<Value>& values) {
  return stddev_helper(values, /*population=*/true);
}

Value AggregateProduct(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  const ArithmeticSummary summary = summarize_arithmetic(values);
  // Excel's pivot PRODUCT on an empty/all-text group returns 0.
  return finite_number(summary.count > 0 ? summary.product : 0.0);
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
Value reify(const Value& v, PivotResult& result) {
  if (!v.is_text()) {
    return v;
  }
  result.text_storage.emplace_back(v.as_text());
  return Value::text(result.text_storage.back());
}
}  // namespace

Value apply_aggregation(Aggregation agg, const std::vector<Value>& values) {
  switch (agg) {
    case Aggregation::Sum:
      return AggregateSum(values);
    case Aggregation::Count:
      return AggregateCount(values);
    case Aggregation::Average:
      return AggregateAverage(values);
    case Aggregation::Max:
      return AggregateMax(values);
    case Aggregation::Min:
      return AggregateMin(values);
    case Aggregation::Product:
      return AggregateProduct(values);
    case Aggregation::CountNumbers:
      return AggregateCountNumbers(values);
    case Aggregation::StdDev:
      return AggregateStdDev(values);
    case Aggregation::StdDevP:
      return AggregateStdDevP(values);
    case Aggregation::Var:
      return AggregateVar(values);
    case Aggregation::VarP:
      return AggregateVarP(values);
  }
  return Value::error(ErrorCode::NA);
}

std::optional<double> numeric_aggregate_value(const Value& v) {
  if (v.is_number()) {
    return v.as_number();
  }
  if (v.is_boolean()) {
    return v.as_boolean() ? 1.0 : 0.0;
  }
  return std::nullopt;
}

void append_record_field_values(const PivotCache& cache, const std::vector<std::size_t>& records,
                                std::uint32_t field_index, std::vector<Value>& out) {
  for (std::size_t rec_idx : records) {
    out.push_back(cell_value(cache, cache.records()[rec_idx], field_index));
  }
}

void append_bucket_field_values(const PivotCache& cache, const RecordBuckets& buckets, std::size_t row_leaf,
                                std::size_t col_leaf, std::uint32_t field_index, std::vector<Value>& out) {
  if (row_leaf >= buckets.size() || col_leaf >= buckets[row_leaf].size()) {
    return;
  }
  append_record_field_values(cache, buckets[row_leaf][col_leaf], field_index, out);
}

void append_leaf_set_field_values(const PivotCache& cache, const RecordBuckets& buckets,
                                  const std::vector<std::size_t>& row_leaves,
                                  const std::vector<std::size_t>& col_leaves, std::uint32_t field_index,
                                  std::vector<Value>& out) {
  for (std::size_t row_leaf : row_leaves) {
    for (std::size_t col_leaf : col_leaves) {
      append_bucket_field_values(cache, buckets, row_leaf, col_leaf, field_index, out);
    }
  }
}

Value aggregate_or_blank(Aggregation aggregation, const std::vector<Value>& values, PivotResult& result) {
  if (values.empty()) {
    return Value::blank();
  }
  return reify(apply_aggregation(aggregation, values), result);
}

}  // namespace formulon::pivot
