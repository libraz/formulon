// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the workday-arithmetic lazy impls: `NETWORKDAYS`,
// `NETWORKDAYS.INTL`, `WORKDAY`, and `WORKDAY.INTL`. See
// `eval/workdays_lazy.h` for the dispatch-table contract and
// `eval/lazy_impls.h` for the shared `eval_node` / `LazyImpl` vocabulary.
//
// The four impls share three helpers: `is_weekend_masked` (generalised
// weekend check parameterised on a 7-bit Mon..Sun mask),
// `parse_weekend_arg` (decodes Excel's numeric / string `weekend`
// selector), and `collect_holidays_from_arg` (reads a holiday set from
// one AST argument node). `is_weekend` (Sat+Sun only) is retained as a
// thin wrapper over `is_weekend_masked(serial, 0x60)` so the original
// NETWORKDAYS / WORKDAY call sites stay unchanged.

#include "eval/workdays_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string_view>
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
// day marked as weekend by `weekend_mask`. The mask is a 7-bit value:
// bit 0 = Monday, ..., bit 6 = Sunday. The standard Sat+Sun weekend is
// `0x60` (bits 5 and 6 set).
//
// The ISO Mon=0 weekday index is derived from the `(days + 4) mod 7`
// Sun=0 form used by WEEKDAY: subtracting one and wrapping gives Mon=0,
// i.e. `(days + 3) mod 7`. 2024-01-01 is a Monday, so `(days(2024,1,1) +
// 3) % 7 == 0` is the canonical cross-check.
bool is_weekend_masked(double serial_floor, std::uint8_t weekend_mask) noexcept {
  const date_time::YMD ymd = date_time::ymd_from_serial(serial_floor);
  const std::int64_t days = date_time::days_from_civil(ymd.y, ymd.m, ymd.d);
  const int mon0 = static_cast<int>(((days + 3) % 7 + 7) % 7);
  return (weekend_mask & (1U << mon0)) != 0U;
}

// Thin wrapper retained so the original NETWORKDAYS / WORKDAY call sites
// keep their self-documenting name. Saturday + Sunday is `0x60` in the
// Mon=0..Sun=6 bit convention.
bool is_weekend(double serial_floor) noexcept {
  return is_weekend_masked(serial_floor, 0x60U);
}

// Decodes the Excel `weekend` argument into a 7-bit Mon=0..Sun=6 mask.
// Accepted shapes:
//
//   * Number 1..7  -> pre-defined paired-weekend pattern (Sat+Sun, Sun+Mon, ...)
//   * Number 11..17 -> single-day weekend (Sun only, Mon only, ...)
//   * 7-char text of '0'/'1' with position 0 = Monday. The all-weekend
//     mask "1111111" is rejected by Excel as `#VALUE!`.
//
// On failure writes the error code into `*out_err` and returns false. The
// error code distinguishes `#NUM!` (invalid numeric selector) from
// `#VALUE!` (malformed string / unsupported kind).
bool parse_weekend_arg(const Value& arg_val, std::uint8_t* out_mask, ErrorCode* out_err) noexcept {
  if (arg_val.is_number()) {
    const double trunc = std::trunc(arg_val.as_number());
    // The table below expands to Excel's documented encoding:
    //   paired:  1 Sat+Sun, 2 Sun+Mon, 3 Mon+Tue, 4 Tue+Wed,
    //            5 Wed+Thu, 6 Thu+Fri, 7 Fri+Sat
    //   single: 11 Sun only, 12 Mon only, 13 Tue only, 14 Wed only,
    //           15 Thu only, 16 Fri only, 17 Sat only
    static constexpr std::uint8_t kPaired[8] = {0U, 0x60U, 0x41U, 0x03U, 0x06U, 0x0CU, 0x18U, 0x30U};
    static constexpr std::uint8_t kSingle[8] = {0U, 0x40U, 0x01U, 0x02U, 0x04U, 0x08U, 0x10U, 0x20U};
    const int n = static_cast<int>(trunc);
    if (n >= 1 && n <= 7) {
      *out_mask = kPaired[n];
      return true;
    }
    if (n >= 11 && n <= 17) {
      *out_mask = kSingle[n - 10];
      return true;
    }
    *out_err = ErrorCode::Num;
    return false;
  }
  if (arg_val.is_text()) {
    const std::string_view s = arg_val.as_text();
    if (s.size() != 7U) {
      *out_err = ErrorCode::Value;
      return false;
    }
    std::uint8_t mask = 0U;
    for (std::size_t i = 0; i < 7U; ++i) {
      const char ch = s[i];
      if (ch == '1') {
        mask |= static_cast<std::uint8_t>(1U << i);
      } else if (ch != '0') {
        *out_err = ErrorCode::Value;
        return false;
      }
    }
    // Excel rejects a mask that marks every day as weekend; there would be
    // no candidate working day to return.
    if (mask == 0x7FU) {
      *out_err = ErrorCode::Value;
      return false;
    }
    *out_mask = mask;
    return true;
  }
  // Blank / Bool / Error / other shapes -- surface #VALUE! consistently.
  *out_err = ErrorCode::Value;
  return false;
}

// Collects holiday serials from a single AST argument node. Accepts four
// shapes:
//
//   * RangeOp / Ref (resolved via resolve_range_arg)
//   * ArrayLiteral  (evaluated cell by cell)
//   * scalar expression (evaluated as a 1-element vector)
//
// Each non-error, non-blank cell is coerced to a number and floored to
// the date component. Blank cells in a range are skipped silently (Excel
// does this to tolerate empty rows inside a holiday column). A
// text / bool / array cell fails the coercion and surfaces as #VALUE!.
// Errors inside the holiday set propagate as the function's result.
//
// On success, fills `out_holidays` (unsorted). On failure, writes the
// error value into `*out_err` and returns false.
bool collect_holidays_from_arg(const parser::AstNode& hol_arg, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, std::vector<double>* out_holidays, Value* out_err) {
  out_holidays->clear();
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

// Thin wrapper preserved for the non-INTL call sites. Routes through
// `collect_holidays_from_arg` when a 3-arg form includes a holidays slot,
// or returns the empty set when the caller's arity is only 2.
bool collect_holidays(const parser::AstNode& call, std::uint32_t arity, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, std::vector<double>* out_holidays, Value* out_err) {
  out_holidays->clear();
  if (arity < 3U) {
    return true;
  }
  return collect_holidays_from_arg(call.as_call_arg(2), arena, registry, ctx, out_holidays, out_err);
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

Value eval_networkdays_intl_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                 const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 4U) {
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
  // Default to the Sat+Sun weekend (matches NETWORKDAYS / selector 1).
  std::uint8_t mask = 0x60U;
  if (arity >= 3U) {
    const Value w = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (w.is_error()) {
      return w;
    }
    ErrorCode ec = ErrorCode::Value;
    if (!parse_weekend_arg(w, &mask, &ec)) {
      return Value::error(ec);
    }
  }
  std::vector<double> holidays;
  Value hol_err = Value::blank();
  if (arity >= 4U) {
    if (!collect_holidays_from_arg(call.as_call_arg(3), arena, registry, ctx, &holidays, &hol_err)) {
      return hol_err;
    }
  }
  std::sort(holidays.begin(), holidays.end());
  holidays.erase(std::unique(holidays.begin(), holidays.end()), holidays.end());

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
    if (is_weekend_masked(d, mask)) {
      continue;
    }
    if (is_holiday_sorted(d, holidays)) {
      continue;
    }
    ++count;
  }
  return Value::number(static_cast<double>(reversed ? -count : count));
}

Value eval_workday_intl_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 4U) {
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
  std::uint8_t mask = 0x60U;
  if (arity >= 3U) {
    const Value w = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (w.is_error()) {
      return w;
    }
    ErrorCode ec = ErrorCode::Value;
    if (!parse_weekend_arg(w, &mask, &ec)) {
      return Value::error(ec);
    }
  }
  std::vector<double> holidays;
  Value hol_err = Value::blank();
  if (arity >= 4U) {
    if (!collect_holidays_from_arg(call.as_call_arg(3), arena, registry, ctx, &holidays, &hol_err)) {
      return hol_err;
    }
  }
  std::sort(holidays.begin(), holidays.end());
  holidays.erase(std::unique(holidays.begin(), holidays.end()), holidays.end());

  double cur = std::floor(start_n.value());
  long long remaining = static_cast<long long>(std::trunc(days_n.value()));
  if (remaining == 0) {
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
    if (is_weekend_masked(cur, mask)) {
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
