//
// Lazy-form routing for the date1904-sensitive calendar builtins. The
// calendar functions that interpret or produce a date serial (DATE, YEAR,
// MONTH, DAY, WEEKDAY, EDATE, EOMONTH, WEEKNUM, ISOWEEKNUM, YEARFRAC,
// DATEDIF, DAYS360, DAYS, DATEVALUE, TODAY, NOW) must honour the workbook's
// `<workbookPr date1904>` epoch, which is only reachable via `EvalContext`.
// The eager `FunctionDef` calling convention (`const Value*`, arity, Arena)
// cannot carry it, so these functions are served through a single lazy
// impl (tree-walker) and a shared `DateEntry` lookup (the VM, which has no
// call AST and therefore reuses the eager-style impl directly).
//
// Time-of-day functions (TIME / HOUR / MINUTE / SECOND / TIMEVALUE) remain
// date1904-independent and eager. DAYS is date-aware and therefore follows
// the lazy route as well, which also supplies its array broadcasting.

#ifndef FORMULON_EVAL_DATETIME_LAZY_H_
#define FORMULON_EVAL_DATETIME_LAZY_H_

#include <cstdint>
#include <string_view>

#include "eval/date_time.h"
#include "value.h"

namespace formulon {

class Arena;

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// A date1904-aware calendar impl. Same shape as an eager builtin plus a
/// trailing `date1904` flag threaded from `EvalContext::date1904()`.
using DateImplFn = Value (*)(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);

/// A clock-reading impl. `NOW` and `TODAY` take no arguments, so what they
/// need from the context is not an argument at all but the wall-clock
/// reading; the argument-shaped `DateImplFn` convention cannot carry it.
/// Giving them their own slot keeps the other fifteen calendar impls'
/// signatures untouched.
using ClockImplFn = Value (*)(const date_time::CivilTime& now, bool date1904);

/// One calendar-family entry: the impl plus its arity bounds (the eager
/// dispatcher's arity guard is replicated by the callers below).
struct DateEntry {
  DateImplFn impl;
  std::uint32_t min_arity;
  std::uint32_t max_arity;
  /// Non-null only for the clock builtins. When set, callers must invoke
  /// this with the context's wall-clock reading rather than calling `impl`,
  /// so a pinned workbook stays deterministic. `impl` remains valid and
  /// reads the host clock, serving callers that have no context to offer.
  ClockImplFn clock_impl = nullptr;
};

/// Returns the calendar entry for `name` (canonical UPPERCASE, future-prefix
/// already stripped) or `nullptr` when `name` is not a date1904-sensitive
/// calendar builtin. Used by the VM to reuse the shared impl with
/// `ctx.date1904()` since the VM has no call AST to drive the lazy path.
const DateEntry* find_date_entry(std::string_view name) noexcept;

/// Invokes a date-aware scalar implementation, lifting array arguments
/// cellwise with Excel 365 broadcasting semantics.
///
/// `now` is the caller's wall-clock reading (`EvalContext::wall_clock()`).
/// It is consumed only by entries carrying a `clock_impl`; the calendar
/// entries ignore it, since they derive everything from their arguments.
Value invoke_date_entry(const DateEntry& entry, const Value* args, std::uint32_t arity, Arena& arena, bool date1904,
                        const date_time::CivilTime& now);

/// Single tree-walker lazy impl covering the whole date1904-sensitive
/// calendar family. Evaluates the call's arguments (scalar-only,
/// left-most error wins), then invokes the matching `DateEntry::impl` with
/// `ctx.date1904()`. Registered in the lazy dispatch table under each
/// calendar name.
Value eval_datetime_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DATETIME_LAZY_H_
