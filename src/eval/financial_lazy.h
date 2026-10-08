//
// Lazy impls for the financial family. `IRR`, `MIRR`, `XIRR`, and
// `XNPV` all live here: each needs the un-flattened AST of its `values`
// argument so a bare Ref, a RangeOp, or an inline ArrayLiteral can be
// walked cell-by-cell and fed into the respective iterative / closed-
// form solution. MIRR additionally takes two trailing scalar rates,
// XIRR / XNPV take a parallel `dates` range that must iterate in
// lockstep with `values`, and all of these arguments must not collide
// with the flattened cash-flow positions — another reason the eager
// `accepts_ranges` path does not fit. The eager path in
// `tree_walker::dispatch_call` would flatten every argument to a single
// `Value` before invoking the impl, erasing the range shape required
// here — which is why these impls ride the lazy dispatch seam alongside
// SUMPRODUCT / INDEX / MATCH / NETWORKDAYS.
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// `eval_node` entry point these impls recurse through.

#ifndef FORMULON_EVAL_FINANCIAL_LAZY_H_
#define FORMULON_EVAL_FINANCIAL_LAZY_H_

#include <cstdint>
#include <string_view>

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `IRR(values, [guess=0.1])` - internal rate of return via Newton-
/// Raphson. `values` must be a range, a single cell reference, or an
/// inline array literal; non-numeric cells inside a range are skipped
/// silently (matching Excel). The sequence must contain at least one
/// positive and one negative cash flow or the call returns `#NUM!`.
/// Iteration uses the classic IRR period-0 convention (first cash flow
/// at time 0), which differs from NPV's period-1 indexing.
Value eval_irr_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx);

/// `MIRR(values, finance_rate, reinvest_rate)` - modified internal rate of
/// return. `values` must resolve to a range, a bare cell reference, or an
/// inline array literal carrying at least one positive and one negative
/// cash flow (else `#DIV/0!`). The closed-form solution mirrors Excel:
///
///   npv_pos = sum over i of max(0, v[i]) / (1 + reinvest_rate)^i
///   npv_neg = sum over i of min(0, v[i]) / (1 + finance_rate)^i
///   MIRR    = (-npv_pos * (1 + reinvest_rate)^(n-1) / npv_neg)^(1/(n-1)) - 1
///
/// Non-numeric cells inside a range argument are silently skipped
/// (matching the range-filter behaviour of SUM / AVERAGE / NPV / IRR).
Value eval_mirr_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

/// `XIRR(values, dates, [guess=0.1])` - internal rate of return for a
/// schedule of cash flows that is NOT necessarily periodic. Both
/// `values` and `dates` must resolve to equal-length ranges (RangeOp,
/// bare Ref, or inline ArrayLiteral); pairs where either cell fails to
/// coerce to a number — or where `values` is Blank — are silently
/// skipped, matching Excel's behaviour for cash-flow ranges. Solves
/// `sum_i values[i] / (1 + rate)^((dates[i] - dates[0]) / 365) == 0`
/// via Newton-Raphson from `guess` (default 0.1). Dates are truncated
/// toward zero before use. The schedule must contain at least one
/// positive and one negative surviving cash flow, at least two cash
/// flows total, and `guess > -1` — otherwise `#NUM!`.
Value eval_xirr_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

/// `XNPV(rate, values, dates)` - net present value for a schedule of
/// cash flows that is not necessarily periodic. Closed-form evaluation
/// of `sum_i values[i] / (1 + rate)^((dates[i] - dates[0]) / 365)`.
/// `rate <= -1` yields `#NUM!`. `values` and `dates` filter identically
/// to XIRR: pairs where either cell fails to coerce (or where `values`
/// is Blank) are dropped before the sum.
Value eval_xnpv_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

// ---------------------------------------------------------------------------
// Date1904-sensitive financial family
// ---------------------------------------------------------------------------
//
// COUPPCD / COUPNCD / COUPNUM / COUPDAYBS / COUPDAYSNC / COUPDAYS, ACCRINT /
// ACCRINTM, DISC / INTRATE / RECEIVED / TBILLPRICE / TBILLYIELD / TBILLEQ,
// PRICE* / YIELD* (regular, discount, maturity, and odd-period variants),
// DURATION / MDURATION, and AMORDEGRC / AMORLINC all decompose a date
// serial (directly, or through `yearfrac_for_basis` / `compute_coupon_dates`
// in `eval/builtins/financial_helpers.h` / `eval/coupon_schedule.h`), so
// they need the workbook's `EvalContext::date1904()` the same way the
// calendar family does (see `eval/datetime_lazy.h`) -- unreachable through
// the eager `FunctionDef` calling convention. They live in this TU rather
// than a dedicated one so the financial family's lazy-dispatch machinery
// stays in one place alongside IRR/MIRR/XIRR/XNPV.

/// A date1904-aware financial-family impl. Same shape as an eager builtin
/// plus a trailing `date1904` flag threaded from `EvalContext::date1904()`.
using FinancialDateImplFn = Value (*)(const Value* args, std::uint32_t arity, Arena& arena, bool date1904);

/// One financial-date-family entry: the impl plus its arity bounds (the
/// eager dispatcher's arity guard is replicated by the callers below).
/// `logical_args` has bit `i` set for a logical parameter (ACCRINT's
/// `calc_method`), which accepts a boolean under the Analysis-ToolPak rule.
struct FinancialDateEntry {
  FinancialDateImplFn impl;
  std::uint32_t min_arity;
  std::uint32_t max_arity;
  std::uint32_t logical_args = 0;
};

/// Returns the financial-date entry for `name` (canonical UPPERCASE) or
/// `nullptr` when `name` does not name one of these functions. Used by the
/// VM to reuse the shared impl with `ctx.date1904()` since the VM has no
/// call AST to drive the lazy path.
const FinancialDateEntry* find_financial_date_entry(std::string_view name) noexcept;

/// Single tree-walker lazy impl covering the whole date1904-sensitive
/// financial family. Evaluates the call's arguments (scalar-only,
/// left-most error wins, matching the eager dispatcher's contract), then
/// invokes the matching `FinancialDateEntry::impl` with `ctx.date1904()`.
/// Registered in the lazy dispatch table under each of the family's names.
Value eval_financial_date_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_FINANCIAL_LAZY_H_
