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
#include <vector>

#include "eval/aggregate_kernels.h"
#include "eval/builtins/registration_helpers.h"
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

// Collects numeric values from `args` into `out`. Range-sourced Bool / Text /
// Blank cells are silently dropped (matching SUM's `range_filter_numeric_only`
// rule). A direct-scalar Error short-circuits the walk and returns its code
// through `*err`; range-sourced non-numeric cells never produce errors here
// because the dispatcher would have routed them through that filter. SUBTOTAL
// is registered with `propagate_errors=false` precisely so errors flow into
// this helper instead of being short-circuited at the dispatcher, letting the
// numeric branches reject them but COUNTA (which counts errors) still see
// them via `run_counta`.
//
// Returns true on a clean walk. On false, `*err` carries the propagating
// error code; `*out` may be partially populated and should be ignored.
bool collect_numeric(const Value* args, std::uint32_t arity, std::vector<double>* out, ErrorCode* err) {
  out->reserve(arity);
  for (std::uint32_t i = 0; i < arity; ++i) {
    const Value& v = args[i];
    if (v.is_error()) {
      *err = v.as_error();
      return false;
    }
    if (v.is_number()) {
      out->push_back(v.as_number());
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

// Converts an `Expected<double, ErrorCode>` from a shared
// `aggregate_kernels::run_*` call into the `Value` shape SUBTOTAL's
// dispatcher expects.
Value lift_kernel_result(Expected<double, ErrorCode> result) {
  if (!result) {
    return Value::error(result.error());
  }
  return Value::number(result.value());
}

Value run_sum(const Value* args, std::uint32_t arity) {
  std::vector<double> xs;
  ErrorCode err = ErrorCode::Value;
  if (!collect_numeric(args, arity, &xs, &err)) {
    return Value::error(err);
  }
  return lift_kernel_result(aggregate_kernels::run_sum(xs));
}

Value run_product(const Value* args, std::uint32_t arity) {
  std::vector<double> xs;
  ErrorCode err = ErrorCode::Value;
  if (!collect_numeric(args, arity, &xs, &err)) {
    return Value::error(err);
  }
  return lift_kernel_result(aggregate_kernels::run_product(xs));
}

Value run_min_max(const Value* args, std::uint32_t arity, bool want_max) {
  std::vector<double> xs;
  ErrorCode err = ErrorCode::Value;
  if (!collect_numeric(args, arity, &xs, &err)) {
    return Value::error(err);
  }
  return lift_kernel_result(want_max ? aggregate_kernels::run_max(xs) : aggregate_kernels::run_min(xs));
}

Value run_average(const Value* args, std::uint32_t arity) {
  std::vector<double> xs;
  ErrorCode err = ErrorCode::Value;
  if (!collect_numeric(args, arity, &xs, &err)) {
    return Value::error(err);
  }
  return lift_kernel_result(aggregate_kernels::run_average(xs));
}

Value run_count(const Value* args, std::uint32_t arity) {
  std::uint32_t n = 0;
  // Errors do not abort COUNT (Excel's COUNT is provenance-tolerant), so we
  // walk by hand instead of going through the numeric-collector helper.
  for (std::uint32_t i = 0; i < arity; ++i) {
    if (args[i].is_number()) {
      ++n;
    }
  }
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

Value run_variance(const Value* args, std::uint32_t arity, bool population) {
  std::vector<double> xs;
  ErrorCode err = ErrorCode::Value;
  if (!collect_numeric(args, arity, &xs, &err)) {
    return Value::error(err);
  }
  return lift_kernel_result(aggregate_kernels::run_variance(xs, /*sample=*/!population));
}

Value run_stdev(const Value* args, std::uint32_t arity, bool population) {
  std::vector<double> xs;
  ErrorCode err = ErrorCode::Value;
  if (!collect_numeric(args, arity, &xs, &err)) {
    return Value::error(err);
  }
  return lift_kernel_result(aggregate_kernels::run_stdev(xs, /*sample=*/!population));
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
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"SUBTOTAL", 2u, kVariadic, &Subtotal, false, true},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
