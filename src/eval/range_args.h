// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Range-argument resolution helper shared by the conditional aggregators
// (`COUNTIF` / `SUMIF` / `AVERAGEIF` / `COUNTIFS` / ...) and the lookup
// family (`VLOOKUP` / `HLOOKUP` / `INDEX` / `MATCH` / `XLOOKUP`). These
// builtins all need to turn a lazily-parsed AST argument — which may be
// either a `RangeOp` (`A1:B2`) or a bare `Ref` — into a flat row-major
// vector of cell `Value`s plus an optional (rows, cols) shape.
//
// Hoisting this helper out of `tree_walker.cpp` lets the conditional and
// lookup impls be split into their own translation units without each TU
// pulling in the rest of the evaluator's internals.

#ifndef FORMULON_EVAL_RANGE_ARGS_H_
#define FORMULON_EVAL_RANGE_ARGS_H_

#include <cstdint>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

class Arena;

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// Resolution result for a range-shaped argument: a flat row-major vector
/// of cell `Value`s plus the rectangle's shape. A 1-cell `Ref` produces
/// `rows = cols = 1`; a `RangeOp` produces the computed shape; a flattened
/// `Value::Array` from a dynamic-array producer carries the array's own
/// dimensions.
struct RangeResult {
  std::vector<Value> cells;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
};

/// Resolves `arg_node` to a `RangeResult`, expanding `RangeOp` / `Ref` /
/// `SpillRef` / range-shaped Calls (`OFFSET` / `CHOOSE` / `IF` / `ROW` /
/// `COLUMN`) and unwrapping dynamic-array producers (`MUNIT`,
/// `SEQUENCE`, lambda invocations, …) into row-major cells.
///
/// Excel-level errors (`#REF!` for a missing sheet, `#VALUE!` for an
/// unsupported argument shape, propagated cell errors, …) are returned
/// as the `Unexpected<ErrorCode>` payload. The Excel `ErrorCode` is the
/// correct error vocabulary here because `resolve_range_arg` is invoked
/// during formula evaluation; engine-level `Error` codes (zip-bomb,
/// arena exhaustion, …) cannot originate here.
///
/// This is the only public surface; the previous `bool + out_param`
/// transitional overload was deleted once the lazy-family fan-out
/// finished migrating.
Expected<RangeResult, ErrorCode> resolve_range_arg(const parser::AstNode& arg_node, Arena& arena,
                                                   const FunctionRegistry& registry, const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_RANGE_ARGS_H_
