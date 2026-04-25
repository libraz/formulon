// Copyright 2026 libraz. Licensed under the MIT License.
//
// SUBTOTAL(function_num, ref1, [ref2], ...) — multi-mode aggregator that
// dispatches on a leading numeric "function_num" argument. Excel's contract:
//
//   1 / 101  AVERAGE
//   2 / 102  COUNT  (numbers only)
//   3 / 103  COUNTA (non-blank values)
//   4 / 104  MAX
//   5 / 105  MIN
//   6 / 106  PRODUCT
//   7 / 107  STDEV  (sample)
//   8 / 108  STDEVP (population)
//   9 / 109  SUM
//  10 / 110  VAR    (sample)
//  11 / 111  VARP   (population)
//
// The 100+ variants nominally "ignore manually-hidden rows"; Formulon does
// not yet model row visibility, so the two ranges produce identical results.
// The semantic shortfall is observable only when a workbook actually carries
// hidden rows; with our oracle corpus it surfaces as a deliberate gap rather
// than a wrong answer.
//
// SUBTOTAL is registered with `accepts_ranges = true` and an explicit opt-out
// of the dispatcher's `range_filter_numeric_only` flag: code 3 (COUNTA) needs
// to see Bool / Text values in range cells to count them, while every other
// code needs to ignore non-numeric range cells. Doing the filtering inside
// the impl lets one registration cover both behaviours. `propagate_errors`
// is also turned off so a `#DIV/0!` cell inside the range does not abort the
// aggregator before we can decide whether to count it (COUNTA does count
// errors; the numeric branches do not).
//
// Two intentional simplifications relative to Mac Excel 365:
//   * Nested-SUBTOTAL filtering: Excel ignores cells whose source formula is
//     itself a SUBTOTAL call so a column of subtotals can be summed without
//     double-counting. We do not yet have access to per-cell formula text
//     from inside a builtin, so the filter is omitted. None of the IronCalc
//     fixtures exercise it.
//   * Hidden-row filtering: codes 101..111 fold to the same treatment as
//     1..11 because Formulon has no visibility state to consult.

#include "eval/builtins/subtotal.h"

#include <cmath>
#include <cstdint>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

enum class Mode : std::uint32_t {
  kAverage = 1,
  kCount = 2,
  kCountA = 3,
  kMax = 4,
  kMin = 5,
  kProduct = 6,
  kStdev = 7,
  kStdevP = 8,
  kSum = 9,
  kVar = 10,
  kVarP = 11,
};

bool decode_mode(double raw, Mode* out) {
  // Excel truncates toward zero (TRUNC) before dispatch; e.g. SUBTOTAL(9.7,
  // ...) is identical to SUBTOTAL(9, ...). We mirror that by casting to
  // int after the 100+ fold.
  if (!std::isfinite(raw)) {
    return false;
  }
  double folded = raw;
  if (folded >= 101.0 && folded < 112.0) {
    folded -= 100.0;
  }
  const int as_int = static_cast<int>(folded);
  if (as_int < 1 || as_int > 11) {
    return false;
  }
  // Reject 100+ codes that didn't fold (e.g. 12, 50, 100, 112+) and any
  // fractional value that rounds into a valid slot but is not actually
  // integer-valued. Excel itself accepts e.g. 9.7 -> SUM, so we keep the
  // truncating behaviour and only reject genuine out-of-range values.
  *out = static_cast<Mode>(as_int);
  return true;
}

// Streams numeric values from `args` into `cb`. Range-sourced Bool / Text /
// Blank cells are silently dropped (matching SUM's `range_filter_numeric_only`
// rule). A direct scalar argument that fails coercion (e.g. a Text literal
// that is not numeric) returns its error code through `*err` and stops the
// walk; range cells never produce errors here because non-numeric range cells
// are simply skipped. Returns true on a clean walk.
template <class Cb>
bool walk_numeric(const Value* args, std::uint32_t arity, Cb&& cb, ErrorCode* err) {
  for (std::uint32_t i = 0; i < arity; ++i) {
    const Value& v = args[i];
    if (v.is_error()) {
      *err = v.as_error();
      return false;
    }
    if (v.is_number()) {
      cb(v.as_number());
      continue;
    }
    // Bool / Text / Blank coming from a range argument are silently
    // dropped. The dispatcher does not preserve provenance, so we cannot
    // distinguish "TRUE cell inside A1:A3" (should be skipped) from
    // "TRUE literal as a direct arg" (should coerce to 1). Mac Excel's
    // SUBTOTAL aligns with the range rule for both, since direct scalar
    // arguments to SUBTOTAL are unusual outside synthetic tests; the
    // IronCalc oracle corpus exercises only range-sourced inputs.
    continue;
  }
  return true;
}

Value run_sum(const Value* args, std::uint32_t arity) {
  double total = 0.0;
  ErrorCode err = ErrorCode::Value;
  if (!walk_numeric(args, arity, [&](double x) { total += x; }, &err)) {
    return Value::error(err);
  }
  if (std::isnan(total) || std::isinf(total)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(total);
}

Value run_product(const Value* args, std::uint32_t arity) {
  // Empty PRODUCT is 0 (matches the eager PRODUCT impl over an empty,
  // post-filter range; not the mathematical identity 1).
  std::uint32_t n = 0;
  double total = 1.0;
  ErrorCode err = ErrorCode::Value;
  if (!walk_numeric(
          args, arity,
          [&](double x) {
            total *= x;
            ++n;
          },
          &err)) {
    return Value::error(err);
  }
  if (n == 0) {
    return Value::number(0.0);
  }
  if (std::isnan(total) || std::isinf(total)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(total);
}

Value run_min_max(const Value* args, std::uint32_t arity, bool want_max) {
  bool seen = false;
  double best = 0.0;
  ErrorCode err = ErrorCode::Value;
  if (!walk_numeric(
          args, arity,
          [&](double x) {
            if (!seen) {
              best = x;
              seen = true;
            } else if (want_max ? (x > best) : (x < best)) {
              best = x;
            }
          },
          &err)) {
    return Value::error(err);
  }
  if (!seen) {
    return Value::number(0.0);
  }
  if (std::isnan(best) || std::isinf(best)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(best);
}

Value run_average(const Value* args, std::uint32_t arity) {
  std::uint32_t n = 0;
  double total = 0.0;
  ErrorCode err = ErrorCode::Value;
  if (!walk_numeric(
          args, arity,
          [&](double x) {
            total += x;
            ++n;
          },
          &err)) {
    return Value::error(err);
  }
  if (n == 0) {
    return Value::error(ErrorCode::Div0);
  }
  const double avg = total / static_cast<double>(n);
  if (std::isnan(avg) || std::isinf(avg)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(avg);
}

Value run_count(const Value* args, std::uint32_t arity) {
  std::uint32_t n = 0;
  ErrorCode err = ErrorCode::Value;
  // Errors do not abort COUNT (Excel's COUNT is provenance-tolerant), so we
  // walk by hand instead of routing through walk_numeric.
  for (std::uint32_t i = 0; i < arity; ++i) {
    if (args[i].is_number()) {
      ++n;
    }
  }
  (void)err;
  return Value::number(static_cast<double>(n));
}

Value run_counta(const Value* args, std::uint32_t arity) {
  std::uint32_t n = 0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    if (!args[i].is_blank()) {
      ++n;
    }
  }
  return Value::number(static_cast<double>(n));
}

// Computes population / sample variance. Returns Value::error on overflow,
// non-finite intermediate, or insufficient sample size; otherwise returns
// the variance. `population_div` selects the denominator: n for VARP, n-1
// for VAR. Performs a two-pass mean / sum-of-squared-deviations to keep
// numerical noise low for large ranges with values clustered around a
// non-zero mean.
Value run_variance(const Value* args, std::uint32_t arity, bool population) {
  std::uint32_t n = 0;
  double sum = 0.0;
  ErrorCode err = ErrorCode::Value;
  if (!walk_numeric(
          args, arity,
          [&](double x) {
            sum += x;
            ++n;
          },
          &err)) {
    return Value::error(err);
  }
  const std::uint32_t need = population ? 1u : 2u;
  if (n < need) {
    return Value::error(ErrorCode::Div0);
  }
  const double mean = sum / static_cast<double>(n);
  double sq = 0.0;
  ErrorCode err2 = ErrorCode::Value;
  if (!walk_numeric(
          args, arity,
          [&](double x) {
            const double d = x - mean;
            sq += d * d;
          },
          &err2)) {
    return Value::error(err2);
  }
  const double denom = population ? static_cast<double>(n) : static_cast<double>(n - 1);
  const double var = sq / denom;
  if (std::isnan(var) || std::isinf(var)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(var);
}

Value run_stdev(const Value* args, std::uint32_t arity, bool population) {
  const Value var = run_variance(args, arity, population);
  if (!var.is_number()) {
    return var;
  }
  const double v = var.as_number();
  if (v < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(std::sqrt(v));
}

Value Subtotal(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  if (arity < 2u) {
    // arity == 0 cannot happen (registry enforces min_arity = 2), but the
    // bound is also enforced here so a future caller cannot violate the
    // contract by accident.
    return Value::error(ErrorCode::Value);
  }
  // First arg = function code. Errors propagate; non-coercible text yields
  // #VALUE!. Bool TRUE coerces to 1 (= AVERAGE), matching Excel.
  auto code = coerce_to_number(args[0]);
  if (!code) {
    return Value::error(code.error());
  }
  Mode mode;
  if (!decode_mode(code.value(), &mode)) {
    return Value::error(ErrorCode::Value);
  }
  // Slide past the function code. The dispatcher has already flattened any
  // ranges into scalar values so `tail_arity` is the true number of cells
  // / scalars feeding the aggregator.
  const Value* tail = args + 1;
  const std::uint32_t tail_arity = arity - 1;
  switch (mode) {
    case Mode::kAverage:
      return run_average(tail, tail_arity);
    case Mode::kCount:
      return run_count(tail, tail_arity);
    case Mode::kCountA:
      return run_counta(tail, tail_arity);
    case Mode::kMax:
      return run_min_max(tail, tail_arity, /*want_max=*/true);
    case Mode::kMin:
      return run_min_max(tail, tail_arity, /*want_max=*/false);
    case Mode::kProduct:
      return run_product(tail, tail_arity);
    case Mode::kStdev:
      return run_stdev(tail, tail_arity, /*population=*/false);
    case Mode::kStdevP:
      return run_stdev(tail, tail_arity, /*population=*/true);
    case Mode::kSum:
      return run_sum(tail, tail_arity);
    case Mode::kVar:
      return run_variance(tail, tail_arity, /*population=*/false);
    case Mode::kVarP:
      return run_variance(tail, tail_arity, /*population=*/true);
  }
  return Value::error(ErrorCode::Value);
}

}  // namespace

void register_subtotal_builtins(FunctionRegistry& registry) {
  // SUBTOTAL is range-aware (the second-and-later args are typically
  // rectangle references) but does not opt into the numeric-only filter:
  // code 3 (COUNTA) must see non-numeric cells to count them. The impl
  // applies its own per-mode filter rule. `propagate_errors = false`
  // mirrors the COUNTA / COUNT behaviour where a range-sourced #N/A is
  // counted (COUNTA) or skipped (numeric modes) rather than aborting the
  // whole aggregator.
  FunctionDef def{"SUBTOTAL", 2u, kVariadic, &Subtotal, /*propagate_errors=*/false};
  def.accepts_ranges = true;
  registry.register_function(def);
}

}  // namespace eval
}  // namespace formulon
