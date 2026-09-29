//
// Registers Excel's text-conversion built-ins (TEXT, VALUE, NUMBERVALUE)
// into a FunctionRegistry. Kept in its own translation unit because the
// three builtins share the format-string engine under
// `eval/text_format/number_format.{h,cpp}` and the date/time parser under
// `eval/date_text_parse.{h,cpp}`.

#ifndef FORMULON_EVAL_BUILTINS_TEXT_FORMAT_H_
#define FORMULON_EVAL_BUILTINS_TEXT_FORMAT_H_

#include <cstdint>

#include "eval/lazy_impls.h"
#include "utils/date_time.h"

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers VALUETOTEXT, ARRAYTOTEXT, NUMBERVALUE, FIXED, DOLLAR, and the
/// rest of the text-conversion family into `registry`. Intended to be
/// invoked from `register_builtins`. TEXT and VALUE are NOT registered
/// here -- see `text_builtin_impl` / `value_builtin_impl` below.
void register_text_format_builtins(FunctionRegistry& registry);

/// TEXT(value, format_text) impl. Not eager-registered: TEXT is
/// date1904-sensitive (date format codes read the workbook epoch), so it is
/// served through the shared `find_date_entry` hook (VM) and the lazy TEXT
/// wrapper (tree-walker), both of which pass `EvalContext::date1904()` here.
Value text_builtin_impl(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);

/// VALUE(text) impl. Not eager-registered: a year-less date_text fallback
/// ("3/15", "3月15日") needs the wall-clock reading for its current-year
/// default, so it is served through the shared `find_date_entry` hook (VM)
/// and the lazy VALUE wrapper (tree-walker), both of which pass
/// `EvalContext::wall_clock()` here. `date1904` is accepted for signature
/// parity with the rest of the `DateEntry` family but unused: VALUE's
/// date-parse fallback has never rebased for the 1904 date system.
Value value_builtin_impl(const Value* args, std::uint32_t arity, Arena& arena, bool date1904,
                         const date_time::CivilTime& now);

/// Host-clock fallback for `value_builtin_impl`, matching the
/// `DateImplFn` signature every plain `DateEntry::impl` uses. Serves a
/// contextless caller (no `EvalContext` to offer a pinned reading).
Value value_builtin_host_clock_impl(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);

/// ARRAYTOTEXT(array, [format]) must preserve the 2-D shape of range and
/// inline-array arguments, so it rides the lazy dispatch path.
Value eval_arraytotext_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_TEXT_FORMAT_H_
