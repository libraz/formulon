// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for Excel's database-aggregation family: `DSUM`, `DCOUNT`,
// `DCOUNTA`, `DAVERAGE`, `DMAX`, `DMIN`, `DPRODUCT`, `DSTDEV`, `DSTDEVP`,
// `DVAR`, `DVARP`, and `DGET`. All twelve share the signature
// `D*(database, field, criteria)` where `database` and `criteria` are
// both rectangular ranges (first row is a header row) and `field` selects
// one column of `database` by 1-based numeric index or by a
// case-insensitive header-name match.
//
// These cannot ride on the eager `accepts_ranges` path: both range
// arguments need to be observed with their (rows, cols) shape intact so
// header-vs-data rows can be distinguished, and the matching logic
// requires seeing each record row atomically rather than as a flattened
// cell stream. Following the pattern used for `*IFS`, every impl is
// referenced from the central dispatch table in `tree_walker.cpp` and
// shares the `LazyImpl` signature declared in `eval/lazy_impls.h`.

#ifndef FORMULON_EVAL_DATABASE_LAZY_H_
#define FORMULON_EVAL_DATABASE_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `DSUM(database, field, criteria)` — sum of numeric values in the
/// selected field column across records that satisfy `criteria`. Empty
/// match set returns `0`. Non-numeric cells in the field column are
/// silently skipped.
Value eval_dsum_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

/// `DCOUNT(database, field, criteria)` — count of numeric cells in the
/// selected field column across matching records. Bool coerces to 0/1 via
/// `coerce_to_number` and is counted. Empty match set returns `0`.
Value eval_dcount_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `DCOUNTA(database, field, criteria)` — count of non-blank cells in the
/// selected field column across matching records. Text, Bool, Number, and
/// Error cells all count. Empty match set returns `0`.
Value eval_dcounta_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);

/// `DAVERAGE(database, field, criteria)` — arithmetic mean of numeric
/// cells in the selected field column across matching records. Empty
/// match set returns `#DIV/0!`.
Value eval_daverage_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);

/// `DMAX(database, field, criteria)` — maximum of numeric cells. Empty
/// match set returns `0` (Excel quirk, matches `MAXIFS`).
Value eval_dmax_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

/// `DMIN(database, field, criteria)` — minimum of numeric cells. Empty
/// match set returns `0` (Excel quirk, matches `MINIFS`).
Value eval_dmin_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

/// `DPRODUCT(database, field, criteria)` — product of numeric cells.
/// Empty match set returns `0` (Excel quirk; mathematically an empty
/// product would be 1, but Excel reports 0).
Value eval_dproduct_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);

/// `DSTDEV(database, field, criteria)` — sample standard deviation
/// (divisor `n - 1`). Requires at least two numeric matches; fewer
/// returns `#DIV/0!`.
Value eval_dstdev_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `DSTDEVP(database, field, criteria)` — population standard deviation
/// (divisor `n`). Empty numeric match set returns `#DIV/0!`.
Value eval_dstdevp_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);

/// `DVAR(database, field, criteria)` — sample variance (divisor
/// `n - 1`). Requires at least two numeric matches; fewer returns
/// `#DIV/0!`.
Value eval_dvar_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

/// `DVARP(database, field, criteria)` — population variance (divisor
/// `n`). Empty numeric match set returns `#DIV/0!`.
Value eval_dvarp_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);

/// `DGET(database, field, criteria)` — the single field-column value
/// from the uniquely matching record. Zero matches returns `#VALUE!`;
/// more than one match returns `#NUM!`.
Value eval_dget_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DATABASE_LAZY_H_
