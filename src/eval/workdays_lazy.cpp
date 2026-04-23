// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the workday-arithmetic lazy impls (`NETWORKDAYS` and
// `WORKDAY`). See `eval/workdays_lazy.h` for the dispatch-table contract
// and `eval/lazy_impls.h` for the shared `eval_node` / `LazyImpl`
// vocabulary.

#include "eval/workdays_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <vector>

#include "eval/coerce.h"
#include "eval/date_time.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/range_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Returns true when `serial_floor` (an integer Excel serial) falls on a
// Saturday or Sunday. Uses the same `(days + 4) mod 7` trick as WEEKDAY:
// Sun=0..Sat=6 in the rotated convention, so weekend = {0, 6}.
bool is_weekend(double serial_floor) noexcept {
  const date_time::YMD ymd = date_time::ymd_from_serial(serial_floor);
  const std::int64_t days = date_time::days_from_civil(ymd.y, ymd.m, ymd.d);
  const int sun0 = static_cast<int>(((days + 4) % 7 + 7) % 7);
  return sun0 == 0 || sun0 == 6;
}

// Collects holiday serials from the third argument of NETWORKDAYS /
// WORKDAY. Accepts four shapes:
//
//   - argument absent (arity == 2)                -> empty set
//   - RangeOp / Ref (resolved via resolve_range_arg)
//   - ArrayLiteral  (evaluated cell by cell)
//   - scalar expression (evaluated as a 1-element vector)
//
// Each non-error, non-blank cell is coerced to a number and floored to
// the date component. Blank cells in a range are skipped silently (Excel
// does this to tolerate empty rows inside a holiday column). A
// text / bool / array cell fails the coercion and surfaces as #VALUE!.
// Errors inside the holiday set propagate as the function's result.
//
// On success, fills `out_holidays` (unsorted). On failure, writes the
// error code into `*out_err` and returns false.
bool collect_holidays(const parser::AstNode& call, std::uint32_t arity, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, std::vector<double>* out_holidays, Value* out_err) {
  out_holidays->clear();
  if (arity < 3U) {
    return true;
  }
  const parser::AstNode& hol_arg = call.as_call_arg(2);
  const parser::NodeKind k = hol_arg.kind();
  std::vector<Value> cells;
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp) {
    ErrorCode err_code = ErrorCode::Value;
    if (!resolve_range_arg(hol_arg, arena, registry, ctx, &cells, &err_code)) {
      *out_err = Value::error(err_code);
      return false;
    }
  } else if (k == parser::NodeKind::ArrayLiteral) {
    const std::uint32_t rows = hol_arg.as_array_rows();
    const std::uint32_t cols = hol_arg.as_array_cols();
    cells.reserve(static_cast<std::size_t>(rows) * cols);
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c) {
        cells.push_back(eval_node(hol_arg.as_array_element(r, c), arena, registry, ctx));
      }
    }
  } else {
    // Scalar argument -- evaluate and treat as a 1-element set. Errors
    // propagate via the `is_error` check below.
    cells.push_back(eval_node(hol_arg, arena, registry, ctx));
  }
  for (const Value& v : cells) {
    if (v.is_error()) {
      *out_err = v;
      return false;
    }
    if (v.is_blank()) {
      // Blank holiday cells are ignored. Matches Excel tolerating empty
      // rows inside a holiday column.
      continue;
    }
    auto n = coerce_to_number(v);
    if (!n) {
      *out_err = Value::error(n.error());
      return false;
    }
    if (n.value() < 0.0) {
      *out_err = Value::error(ErrorCode::Num);
      return false;
    }
    out_holidays->push_back(std::floor(n.value()));
  }
  return true;
}

// Returns true when `day_serial` (integer serial) is present in the
// pre-sorted `holidays` vector. Caller must sort once before the loop.
bool is_holiday_sorted(double day_serial, const std::vector<double>& holidays) noexcept {
  return std::binary_search(holidays.begin(), holidays.end(), day_serial);
}

}  // namespace

Value eval_networkdays_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }
  const Value start_v = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (start_v.is_error()) {
    return start_v;
  }
  const Value end_v = eval_node(call.as_call_arg(1), arena, registry, ctx);
  if (end_v.is_error()) {
    return end_v;
  }
  auto start_n = coerce_to_number(start_v);
  if (!start_n) {
    return Value::error(start_n.error());
  }
  auto end_n = coerce_to_number(end_v);
  if (!end_n) {
    return Value::error(end_n.error());
  }
  if (start_n.value() < 0.0 || end_n.value() < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  std::vector<double> holidays;
  Value hol_err = Value::blank();
  if (!collect_holidays(call, arity, arena, registry, ctx, &holidays, &hol_err)) {
    return hol_err;
  }
  std::sort(holidays.begin(), holidays.end());
  holidays.erase(std::unique(holidays.begin(), holidays.end()), holidays.end());

  // Walk the interval in either direction. Excel 365 returns a negative
  // count when `start > end`, matching the oracle's behaviour.
  double s = std::floor(start_n.value());
  double e = std::floor(end_n.value());
  const bool reversed = s > e;
  if (reversed) {
    const double tmp = s;
    s = e;
    e = tmp;
  }
  long long count = 0;
  for (double d = s; d <= e; d += 1.0) {
    if (is_weekend(d)) {
      continue;
    }
    if (is_holiday_sorted(d, holidays)) {
      continue;
    }
    ++count;
  }
  return Value::number(static_cast<double>(reversed ? -count : count));
}

Value eval_workday_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }
  const Value start_v = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (start_v.is_error()) {
    return start_v;
  }
  const Value days_v = eval_node(call.as_call_arg(1), arena, registry, ctx);
  if (days_v.is_error()) {
    return days_v;
  }
  auto start_n = coerce_to_number(start_v);
  if (!start_n) {
    return Value::error(start_n.error());
  }
  auto days_n = coerce_to_number(days_v);
  if (!days_n) {
    return Value::error(days_n.error());
  }
  if (start_n.value() < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  std::vector<double> holidays;
  Value hol_err = Value::blank();
  if (!collect_holidays(call, arity, arena, registry, ctx, &holidays, &hol_err)) {
    return hol_err;
  }
  std::sort(holidays.begin(), holidays.end());
  holidays.erase(std::unique(holidays.begin(), holidays.end()), holidays.end());

  double cur = std::floor(start_n.value());
  long long remaining = static_cast<long long>(std::trunc(days_n.value()));
  if (remaining == 0) {
    // Excel WORKDAY(start, 0) returns start unchanged (no weekend/holiday
    // adjustment). This is the canonical behaviour confirmed by 365.
    return Value::number(cur);
  }
  const int step = remaining > 0 ? 1 : -1;
  if (remaining < 0) {
    remaining = -remaining;
  }
  while (remaining > 0) {
    cur += step;
    if (cur < 0.0) {
      return Value::error(ErrorCode::Num);
    }
    if (is_weekend(cur)) {
      continue;
    }
    if (is_holiday_sorted(cur, holidays)) {
      continue;
    }
    --remaining;
  }
  return Value::number(cur);
}

}  // namespace eval
}  // namespace formulon
