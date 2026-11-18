//
// Implementation of the workday-arithmetic lazy impls: `NETWORKDAYS`,
// `NETWORKDAYS.INTL`, `WORKDAY`, and `WORKDAY.INTL`. See
// `eval/workdays_lazy.h` for the dispatch-table contract and
// `eval/lazy_impls.h` for the shared `eval_node` / `LazyImpl` vocabulary.
//
// The non-INTL forms are the INTL forms with the fixed Sat+Sun mask and
// the holiday list in slot 2: all four share the leading-argument reader
// and the `count_workdays` / `add_workdays` walks, parameterised on a
// 7-bit Mon..Sun weekend mask.

#include "eval/workdays_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/omitted_arg.h"
#include "eval/range_args.h"
#include "eval/scalar_args.h"
#include "eval/tree_walker/dispatch.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Upper bound of the Excel serial range, 9999-12-31. Under the 1904 date
// system every serial denotes a calendar day 1462 days later, so the
// bound shifts with the epoch rather than being a single constant.
double max_workday_serial(bool date1904) noexcept {
  return date_time::serial_from_ymd(9999, 12, 31, date1904);
}

bool is_valid_workday_serial(double value, bool date1904) noexcept {
  return std::isfinite(value) && value >= 0.0 && value <= max_workday_serial(date1904);
}

bool is_valid_workday_count(double value, bool date1904) noexcept {
  return std::isfinite(value) && std::fabs(std::trunc(value)) <= max_workday_serial(date1904);
}

// Returns true when `serial_floor` (an integer Excel serial, interpreted
// under the workbook's `date1904` epoch) falls on a day marked as
// weekend by `weekend_mask`. The mask is a 7-bit value: bit 0 = Monday,
// ..., bit 6 = Sunday. The standard Sat+Sun weekend is `0x60` (bits 5
// and 6 set).
//
// The epoch matters: the 1900/1904 offset of 1462 days is not a multiple
// of 7, so reading a 1904 serial as a 1900 one shifts the weekday by two
// days and silently changes every weekend decision.
bool is_weekend_masked(double serial_floor, std::uint8_t weekend_mask, bool date1904) noexcept {
  const int mon0 = (date_time::weekday_sun0(serial_floor, date1904) + 6) % 7;
  return (weekend_mask & (1U << mon0)) != 0U;
}

// Saturday + Sunday in the Mon=0..Sun=6 bit convention (selector 1).
constexpr std::uint8_t kSatSunWeekendMask = 0x60U;

// Decodes the Excel `weekend` argument into a 7-bit Mon=0..Sun=6 mask.
// Accepted shapes:
//
//   * Number 1..7  -> pre-defined paired-weekend pattern (TRUE reads as 1) (Sat+Sun, Sun+Mon, ...)
//   * Number 11..17 -> single-day weekend (Sun only, Mon only, ...)
//   * 7-char text of '0'/'1' with position 0 = Monday. The all-weekend
//     mask "1111111" is `#VALUE!` for WORKDAY.INTL and counts 0 in NETWORKDAYS.INTL.
//
// On failure writes the error code into `*out_err` and returns false. The
// error code distinguishes `#NUM!` (invalid numeric selector) from
// `#VALUE!` (malformed string / unsupported kind).
bool parse_weekend_arg(const Value& arg_val, bool reject_all_weekend, std::uint8_t* out_mask,
                       ErrorCode* out_err) noexcept {
  if (arg_val.is_number() || arg_val.is_boolean() || arg_val.is_blank()) {
    // TRUE is selector 1; FALSE and a blank cell are selector 0, which no pattern matches.
    const double selector = arg_val.is_number() ? arg_val.as_number() : (arg_val.is_boolean() && arg_val.as_boolean());
    const double trunc = std::trunc(selector);
    if (!std::isfinite(trunc)) {
      *out_err = ErrorCode::Num;
      return false;
    }
    // The table below expands to Excel's documented encoding:
    //   paired:  1 Sat+Sun, 2 Sun+Mon, 3 Mon+Tue, 4 Tue+Wed,
    //            5 Wed+Thu, 6 Thu+Fri, 7 Fri+Sat
    //   single: 11 Sun only, 12 Mon only, 13 Tue only, 14 Wed only,
    //           15 Thu only, 16 Fri only, 17 Sat only
    static constexpr std::uint8_t kPaired[8] = {0U, 0x60U, 0x41U, 0x03U, 0x06U, 0x0CU, 0x18U, 0x30U};
    static constexpr std::uint8_t kSingle[8] = {0U, 0x40U, 0x01U, 0x02U, 0x04U, 0x08U, 0x10U, 0x20U};
    if (trunc >= 1.0 && trunc <= 7.0) {
      const int n = static_cast<int>(trunc);
      *out_mask = kPaired[n];
      return true;
    }
    if (trunc >= 11.0 && trunc <= 17.0) {
      const int n = static_cast<int>(trunc);
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
    // WORKDAY.INTL rejects a mask that marks every day as weekend (no candidate
    // working day); NETWORKDAYS.INTL accepts it and counts 0.
    if (reject_all_weekend && mask == 0x7FU) {
      *out_err = ErrorCode::Value;
      return false;
    }
    *out_mask = mask;
    return true;
  }
  // Error / other shapes -- surface #VALUE! consistently.
  *out_err = ErrorCode::Value;
  return false;
}

// Collects holiday serials from a single AST argument node. Every shape
// `resolve_range_arg` resolves is accepted — `Ref` / `RangeOp` /
// `SpillRef` / `ArrayLiteral` and dynamic-array producers such as
// `SEQUENCE` — and a bare scalar collapses to a 1-element set.
//
// Each non-error, non-blank cell is coerced to a number and floored to
// the date component. Blank cells in a range are skipped silently (Excel
// does this to tolerate empty rows inside a holiday column). A
// text / bool / array cell fails the coercion and surfaces as #VALUE!.
// Errors inside the holiday set propagate as the function's result.
//
// On success, fills `out_holidays` sorted and deduplicated, ready for
// `is_holiday_sorted`. On failure, writes the error value into `*out_err`
// and returns false.
bool collect_holidays_from_arg(const parser::AstNode& hol_arg, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, std::vector<double>* out_holidays, Value* out_err) {
  out_holidays->clear();
  auto resolved = resolve_range_arg(hol_arg, arena, registry, ctx);
  if (!resolved) {
    *out_err = Value::error(resolved.error());
    return false;
  }
  const std::vector<Value> cells = std::move(resolved.value().cells);
  // A lone boolean holiday (literal or cell) is #VALUE! under the Analysis-ToolPak rule.
  if (cells.size() == 1U && cells[0].is_boolean()) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
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
  std::sort(out_holidays->begin(), out_holidays->end());
  out_holidays->erase(std::unique(out_holidays->begin(), out_holidays->end()), out_holidays->end());
  return true;
}

// Returns true when `day_serial` (integer serial) is present in the
// pre-sorted `holidays` vector. Caller must sort once before the loop.
bool is_holiday_sorted(double day_serial, const std::vector<double>& holidays) noexcept {
  return std::binary_search(holidays.begin(), holidays.end(), day_serial);
}

// Counts working days in the closed interval between two validated
// serials. Excel 365 returns a negative count when `start > end`.
Value count_workdays(double start, double end, std::uint8_t mask, const std::vector<double>& holidays, bool date1904) {
  double s = std::floor(start);
  double e = std::floor(end);
  const bool reversed = s > e;
  if (reversed) {
    const double tmp = s;
    s = e;
    e = tmp;
  }
  long long count = 0;
  for (double d = s; d <= e; d += 1.0) {
    if (is_weekend_masked(d, mask, date1904)) {
      continue;
    }
    if (is_holiday_sorted(d, holidays)) {
      continue;
    }
    ++count;
  }
  return Value::number(static_cast<double>(reversed ? -count : count));
}

// Steps `days` working days from a validated start serial; #NUM! once the
// walk leaves the serial range.
Value add_workdays(double start, double days, std::uint8_t mask, const std::vector<double>& holidays, bool date1904) {
  double cur = std::floor(start);
  long long remaining = static_cast<long long>(std::floor(days));
  if (remaining == 0) {
    // Excel WORKDAY(start, 0) returns start unchanged (no weekend/holiday
    // adjustment). This is the canonical behaviour confirmed by 365.
    return Value::number(cur);
  }
  const int step = remaining > 0 ? 1 : -1;
  if (remaining < 0) {
    remaining = -remaining;
  }
  const double max_serial = max_workday_serial(date1904);
  while (remaining > 0) {
    cur += step;
    if (cur < 0.0 || cur > max_serial) {
      return Value::error(ErrorCode::Num);
    }
    if (is_weekend_masked(cur, mask, date1904)) {
      continue;
    }
    if (is_holiday_sorted(cur, holidays)) {
      continue;
    }
    --remaining;
  }
  return Value::number(cur);
}

// One call of NETWORKDAYS / WORKDAY or their .INTL forms over evaluated arguments.
struct WorkdayCall {
  bool add_days;
  bool date1904;
  const std::vector<double>* holidays;
  const Value* holiday_error;  // the holiday argument's error, or blank
};

// An error weekend, alone or as an element, reads as an invalid selector: #NUM! for WORKDAY.INTL, #VALUE! for
// NETWORKDAYS.INTL. Each error becomes a weekend value `parse_weekend_arg` rejects with that code.
Value invalid_weekend_for_errors(const Value& weekend, bool add_days, Arena& arena) {
  const Value invalid = add_days ? Value::number(0.0) : Value::text("");
  if (weekend.is_error()) {
    return invalid;
  }
  if (!weekend.is_array()) {
    return weekend;
  }
  const std::size_t count = static_cast<std::size_t>(weekend.as_array_rows()) * weekend.as_array_cols();
  const Value* cells = weekend.as_array_cells();
  if (std::none_of(cells, cells + count, [](const Value& v) { return v.is_error(); })) {
    return weekend;
  }
  Value* buf = nullptr;
  ArrayValue* out = allocate_derived_array_value(weekend.as_array_rows(), weekend.as_array_cols(), arena, buf);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < count; ++i) {
    buf[i] = cells[i].is_error() ? invalid : cells[i];
  }
  return Value::array(out);
}

// `a[0]` is the start, `a[1]` the day count or end date and, for the .INTL forms, `a[2]` the weekend.
Value workday_cell(const WorkdayCall& c, const Value* a, std::uint32_t n) {
  auto first_n = coerce_to_number(a[0]);
  if (!first_n) {
    return Value::error(first_n.error());
  }
  auto second_n = coerce_to_number(a[1]);
  if (!second_n) {
    return Value::error(second_n.error());
  }
  // WORKDAY truncates a fractional start before the range check.
  const double start = c.add_days ? std::trunc(first_n.value()) : first_n.value();
  const double second = second_n.value();
  if (!is_valid_workday_serial(start, c.date1904) ||
      !(c.add_days ? is_valid_workday_count(second, c.date1904) : is_valid_workday_serial(second, c.date1904))) {
    return Value::error(ErrorCode::Num);
  }
  std::uint8_t mask = kSatSunWeekendMask;
  if (n > 2U) {
    ErrorCode code = ErrorCode::Value;
    if (!parse_weekend_arg(a[2], c.add_days, &mask, &code)) {
      return Value::error(code);
    }
  }
  if (c.holiday_error->is_error()) {
    return *c.holiday_error;
  }
  return c.add_days ? add_workdays(start, second, mask, *c.holidays, c.date1904)
                    : count_workdays(start, second, mask, *c.holidays, c.date1904);
}

// Shared driver for NETWORKDAYS / WORKDAY and their .INTL forms: `add_days` selects WORKDAY, `intl` the
// weekend-mask calendar (and the 4th argument). An array start, day count, end date or weekend lifts the call;
// a multi-cell reference lifts only as the weekend and is #VALUE! anywhere else.
Value eval_workdays_driver(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx, bool add_days, bool intl) {
  if (call.as_call_arity() < 2U || call.as_call_arity() > (intl ? 4U : 3U)) {
    return Value::error(ErrorCode::Value);
  }
  // Trailing omitted optional slots take their defaults.
  const std::uint32_t arity = atp_evaluated_arity(call, 2U, intl ? 4U : 3U);
  Value args[3] = {Value::blank(), Value::blank(), Value::blank()};
  // An error argument is held, not returned: beside a lifting argument it fills every element.
  const Value* first_error = nullptr;
  for (std::uint32_t i = 0; i < 2U; ++i) {
    const parser::AstNode& arg = call.as_call_arg(i);
    args[i] = eval_node(arg, arena, registry, ctx);
    if (args[i].is_error()) {
      first_error = first_error != nullptr ? first_error : &args[i];
      continue;
    }
    if (args[i].is_array() && is_reference_shape(arg)) {
      return first_error != nullptr ? *first_error : Value::error(ErrorCode::Value);
    }
    if (const Value atp = atp_arg_error(arg, args[i], /*required=*/true); atp.is_error()) {
      args[i] = atp;
      first_error = first_error != nullptr ? first_error : &args[i];
    }
  }
  std::uint32_t lifted = 2U;
  if (intl && arity >= 3U && !is_omitted_arg(call.as_call_arg(2))) {
    args[2] = invalid_weekend_for_errors(eval_node(call.as_call_arg(2), arena, registry, ctx), add_days, arena);
    lifted = 3U;
  }
  std::vector<double> holidays;
  Value holiday_error = Value::blank();
  const std::uint32_t holiday_slot = intl ? 3U : 2U;
  if (arity > holiday_slot) {
    collect_holidays_from_arg(call.as_call_arg(holiday_slot), arena, registry, ctx, &holidays, &holiday_error);
  }
  const WorkdayCall c{add_days, ctx.date1904(), &holidays, &holiday_error};
  return lift_over_array_args(args, lifted, arena,
                              [&c](const Value* a, std::uint32_t n) { return workday_cell(c, a, n); });
}

}  // namespace

Value eval_networkdays_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx) {
  return eval_workdays_driver(call, arena, registry, ctx, /*add_days=*/false, /*intl=*/false);
}

Value eval_workday_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  return eval_workdays_driver(call, arena, registry, ctx, /*add_days=*/true, /*intl=*/false);
}

Value eval_networkdays_intl_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                 const EvalContext& ctx) {
  return eval_workdays_driver(call, arena, registry, ctx, /*add_days=*/false, /*intl=*/true);
}

Value eval_workday_intl_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  return eval_workdays_driver(call, arena, registry, ctx, /*add_days=*/true, /*intl=*/true);
}

}  // namespace eval
}  // namespace formulon
