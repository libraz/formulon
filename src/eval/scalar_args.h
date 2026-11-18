// Scalar argument readers shared by the lazy function families.

#ifndef FORMULON_EVAL_SCALAR_ARGS_H_
#define FORMULON_EVAL_SCALAR_ARGS_H_

#include <cstdint>
#include <optional>
#include <vector>

#include "eval/error_origin.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

class Arena;

namespace eval {

class EvalContext;
class FunctionRegistry;

/// Coerce an evaluated value to a number. An error value or a coercion
/// failure is written to `error_out` and yields `false`.
bool number_from_value(const Value& v, double& out, Value& error_out);

/// Coerce an evaluated value through `coerce_to_bool`. An error value or a
/// coercion failure is written to `error_out` and yields `false`.
bool bool_from_value(const Value& v, bool& out, Value& error_out);

/// Evaluate `node` and coerce it to a number. On error or coercion failure
/// writes the caller-visible error to `error_out` and returns `false`.
bool eval_number_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, double& out, Value& error_out);

/// Evaluate `node` and coerce it through `coerce_to_bool`. On error or
/// coercion failure writes the caller-visible error to `error_out` and
/// returns `false`.
bool eval_bool_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                   bool& out, Value& error_out);

/// Evaluate `node`, coerce to a finite number, and truncate toward zero.
/// On error or coercion failure writes the caller-visible error to
/// `error_out` and returns `false`. Shared by every helper that takes a
/// numeric scalar argument (counts, indices, ignore masks, etc.).
bool eval_truncated_number_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, double& out, Value& error_out);

/// Evaluates the scalar arguments of a call with `min_arity`..`max_arity`
/// left to right into `out`, an error argument kept in its slot so the impl's
/// coercion order decides what is reported. With `atp_rule` the
/// Analysis-ToolPak rule applies: trailing omitted optional slots are dropped,
/// and each slot rejects a boolean unless bit `i` of `logical_args` is set
/// (see `atp_arg_error`); a rejected slot holds the error it answers.
/// `first_error_origin` receives the error origin as the first error argument
/// left it.
void eval_scalar_call_args(const parser::AstNode& call, std::uint32_t min_arity, std::uint32_t max_arity, bool atp_rule,
                           std::uint32_t logical_args, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx, std::vector<Value>& out,
                           std::optional<ErrorOriginSnapshot>& first_error_origin);

/// The error argument a propagating function reports when its impl fails on
/// none: the first error, the last with `last_first`, or null.
const Value* reported_error_arg(const Value* args, std::uint32_t arity, bool last_first);

/// `result` of an impl over `args`, or the reported error argument when the
/// impl returned a non-error beside one.
Value report_error_arg(const Value& result, const Value* args, std::uint32_t arity, bool last_first);

/// Keeps the origin of a reported error argument once it is the result: the
/// impl's own failure keeps what it noted. Call with the origin forgotten
/// before the impl ran.
void restore_reported_origin(const Value& result, const Value* args, std::uint32_t arity, bool last_first,
                             const std::optional<ErrorOriginSnapshot>& reported_origin);

/// One scalar evaluation of a lifted call over `arity` per-element arguments.
using LiftedCellFn = Value (*)(const Value* args, std::uint32_t arity, const void* state);

/// A multi-cell reference in a scalar parameter that does not lift: it takes no
/// part in the result's extent, and at result position k (row-major) it gives
/// #VALUE! (k = 0), its first cell (k = 1) or its last cell (k >= 2).
struct AreaEnds {
  bool active = false;
  Value first;
  Value last;
};

/// The extent of a lift over `args`: the widest array or grid argument on each
/// axis, an argument with an active `ends` entry left out (`ends` may be null).
/// False when no argument lifts.
bool lift_extent(const Value* args, std::uint32_t arity, const AreaEnds* ends, std::uint32_t* rows,
                 std::uint32_t* cols);

/// Runs `cell` at each position of a `rows` x `cols` lift over `args`, writing
/// row-major into `out`. An array or grid argument gives its element there (a
/// size-1 axis stretches; a grid element is a one-area `Ref`, an error element
/// its error), an active `ends` entry its `AreaEnds` value, any other argument
/// itself. A position some argument cannot supply is #N/A, and a position whose
/// arguments hold an error takes the left-most one without running `cell`. A
/// position whose `cell` answers an array holds that array's top-left element.
void lift_positions(const Value* args, std::uint32_t arity, const AreaEnds* ends, std::uint32_t rows,
                    std::uint32_t cols, Arena& arena, LiftedCellFn cell, const void* state, Value* out);

/// Applies a scalar function element by element when an argument is an array
/// (`lift_extent`, `lift_positions`), collecting the results into an array.
/// Without a lifting argument `cell` runs once on `args`, unless an `ends`
/// entry is active, which makes the call #VALUE!.
Value lift_over_array_args(const Value* args, std::uint32_t arity, Arena& arena, LiftedCellFn cell, const void* state,
                           const AreaEnds* ends = nullptr);

/// `lift_over_array_args` with a const callable `cell(const Value*, std::uint32_t)`.
template <typename Cell>
Value lift_over_array_args(const Value* args, std::uint32_t arity, Arena& arena, const Cell& cell) {
  return lift_over_array_args(
      args, arity, arena,
      [](const Value* a, std::uint32_t n, const void* state) { return (*static_cast<const Cell*>(state))(a, n); },
      static_cast<const void*>(&cell));
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SCALAR_ARGS_H_
