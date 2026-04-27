// Copyright 2026 libraz. Licensed under the MIT License.
//
// Stateless, header-only Excel scalar-operator primitives: cross-type
// comparison, unary +/-/%, the five arithmetic operators (+, -, *, /, ^),
// the `&` text-concat operator, and the six relational operators
// (=, <>, <, <=, >, >=).
//
// These helpers are shared between the eager scalar dispatch in
// `tree_walker.cpp` and the array-context cellwise broadcaster in
// `shape_ops_lazy.cpp`. Their behaviour is intentionally identical to the
// pre-extraction code that lived in `tree_walker.cpp`'s anonymous namespace
// — keeping a single source of truth means the scalar and array-broadcast
// paths cannot diverge on edge cases (NaN ordering, blank-vs-empty-string
// equality, `0/0 -> #DIV/0!`, etc.).
//
// Strictly scalar: none of these accept `Value::Array`. Broadcasting is
// the caller's responsibility — see `shape_ops_lazy.cpp` for how cellwise
// fan-out is layered on top of these primitives.

#ifndef FORMULON_EVAL_SCALAR_OPS_H_
#define FORMULON_EVAL_SCALAR_OPS_H_

#include "parser/ast.h"
#include "value.h"

namespace formulon {

class Arena;

namespace eval {

/// Excel cross-type comparison. Returns -1 / 0 / +1 with Excel's ordering
/// (Number < Text < Bool; Blank coerces to numeric zero, except in the
/// chameleonic `Blank == ""` case where it compares as the empty string).
/// Sets `*out_unordered` to true iff one of the operands is NaN; in that
/// case the integer result is meaningless and the caller must short-circuit
/// every relational operator to FALSE except `<>`.
int compare_values(const Value& lhs, const Value& rhs, bool* out_unordered);

/// Wraps a finite arithmetic result as a `Number`; `NaN` or `Inf` becomes
/// `#NUM!`.
Value finalize_arithmetic(double r);

/// Applies one of Excel's unary operators (`+`, `-`, `%`). Unary `+` is an
/// identity (does NOT coerce, matching Excel 365); `-` and `%` coerce to
/// number first and surface coercion errors.
Value apply_unary(parser::UnaryOp op, const Value& operand);

/// Applies one of the five arithmetic binary operators (`+`, `-`, `*`, `/`,
/// `^`) on two already-coerced doubles. `0` divisor surfaces `#DIV/0!`;
/// `^` delegates to `apply_pow` so `^` and the `POWER()` builtin cannot
/// drift apart on edge cases.
Value apply_arithmetic(parser::BinOp op, double lhs, double rhs);

/// Applies the `&` text-concatenation operator. Both operands are coerced
/// to text via `coerce_to_text`; the joined string is interned in `arena`
/// so the returned `Text` view remains valid for the arena's lifetime.
Value apply_concat(const Value& lhs, const Value& rhs, Arena& arena);

/// Applies one of the six relational operators (`=`, `<>`, `<`, `<=`, `>`,
/// `>=`) using `compare_values` ordering. NaN unordered handling follows
/// IEEE-754 (every comparison is FALSE except `<>`).
Value apply_comparison(parser::BinOp op, const Value& lhs, const Value& rhs);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SCALAR_OPS_H_
