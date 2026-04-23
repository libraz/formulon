// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's calendar built-ins (DATE, TIME, YEAR, MONTH, DAY, HOUR,
// MINUTE, SECOND, WEEKDAY, EDATE, EOMONTH, DAYS) into a FunctionRegistry.
// Kept in its own translation unit so the date/time family can evolve
// independently of the rest of the builtin catalog. Shared calendar math
// lives in `eval/date_time.{h,cpp}`.

#ifndef FORMULON_EVAL_BUILTINS_DATETIME_H_
#define FORMULON_EVAL_BUILTINS_DATETIME_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the date/time built-in functions (DATE, TIME, YEAR, MONTH, DAY,
/// HOUR, MINUTE, SECOND, WEEKDAY, EDATE, EOMONTH, DAYS) into `registry`.
/// Intended to be invoked from `register_builtins`.
void register_datetime_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_DATETIME_H_
