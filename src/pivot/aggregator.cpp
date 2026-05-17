// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

#include "pivot/aggregator.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <vector>

#include "pivot/pivot_cache.h"
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
  std::vector<double> samples;
};

ArithmeticSummary summarize_arithmetic(const std::vector<Value>& values, bool keep_samples = false) {
  ArithmeticSummary summary;
  if (keep_samples) {
    summary.samples.reserve(values.size());
  }
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
    if (keep_samples) {
      summary.samples.push_back(x);
    }
  }
  return summary;
}

Value AggregateSum(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  return Value::number(summarize_arithmetic(values).sum);
}

// Excel's pivot `Count` mirrors COUNTA: any non-blank cell counts,
// including text and booleans.
Value AggregateCount(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  double count = 0.0;
  for (const auto& v : values) {
    if (!v.is_blank()) {
      count += 1.0;
    }
  }
  return Value::number(count);
}

// Excel's pivot `CountNumbers` mirrors COUNT: only numeric cells
// (booleans included, per Excel).
Value AggregateCountNumbers(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
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
  const ArithmeticSummary summary = summarize_arithmetic(values);
  if (summary.count == 0) {
    return Value::error(ErrorCode::Div0);
  }
  return Value::number(summary.sum / static_cast<double>(summary.count));
}

Value AggregateMax(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  const ArithmeticSummary summary = summarize_arithmetic(values);
  // Excel's pivot MAX over an empty/all-text group returns 0.
  return Value::number(summary.count > 0 ? summary.max : 0.0);
}

Value AggregateMin(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  const ArithmeticSummary summary = summarize_arithmetic(values);
  return Value::number(summary.count > 0 ? summary.min : 0.0);
}

// Two-pass variance computation. Mirrors `VAR.S` / `VAR.P` from
// `src/eval/builtins/stats.cpp`: collect numerics (booleans coerce as
// 0/1, text and blanks are ignored), then compute mean and sum of
// squared deviations. `population` controls the divisor: when true, n;
// otherwise n - 1. The `min_n` guard rejects samples too small for the
// chosen variant (n < 2 for sample, n < 1 for population) with
// `#DIV/0!`, matching Excel's pivot behaviour. Errors in the input
// dominate (handled by the caller's `first_error` short-circuit).
Value variance_helper(const std::vector<Value>& values, bool population) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  const ArithmeticSummary summary = summarize_arithmetic(values, /*keep_samples=*/true);
  const std::vector<double>& xs = summary.samples;
  const std::size_t min_n = population ? 1u : 2u;
  if (xs.size() < min_n) {
    return Value::error(ErrorCode::Div0);
  }
  const double mean = summary.sum / static_cast<double>(xs.size());
  double ss = 0.0;
  for (double x : xs) {
    const double d = x - mean;
    ss += d * d;
  }
  const double divisor = population ? static_cast<double>(xs.size()) : static_cast<double>(xs.size() - 1u);
  const double r = ss / divisor;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

Value AggregateVar(const std::vector<Value>& values) {
  return variance_helper(values, /*population=*/false);
}

Value AggregateVarP(const std::vector<Value>& values) {
  return variance_helper(values, /*population=*/true);
}

Value AggregateStdDev(const std::vector<Value>& values) {
  Value v = variance_helper(values, /*population=*/false);
  if (!v.is_number()) {
    return v;
  }
  const double r = std::sqrt(v.as_number());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

Value AggregateStdDevP(const std::vector<Value>& values) {
  Value v = variance_helper(values, /*population=*/true);
  if (!v.is_number()) {
    return v;
  }
  const double r = std::sqrt(v.as_number());
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

Value AggregateProduct(const std::vector<Value>& values) {
  if (const Value* err = first_error(values); err != nullptr) {
    return *err;
  }
  const ArithmeticSummary summary = summarize_arithmetic(values);
  // Excel's pivot PRODUCT on an empty/all-text group returns 0.
  return Value::number(summary.count > 0 ? summary.product : 0.0);
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

}  // namespace formulon::pivot
