// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Lazy impl for GETPIVOTDATA. The function takes a data-field name + a
// pivot-anchor cell reference + zero or more (field, item) pairs, and
// returns the value of the addressed cell inside the freshest pivot
// evaluation snapshot.
//
// Why lazy: the second argument is required to be a cell or range
// reference -- the eager dispatcher would flatten it to a Value before
// the impl could see the un-evaluated AST and recover the (sheet, row,
// col) coordinates needed to identify which pivot table the caller is
// addressing. The remaining arguments are scalar text/number lookups
// and are evaluated eagerly inside the impl.
//
// On a successful lookup the impl returns the cached `PivotResult`
// value at the resolved leaf; on any failure (anchor not over a pivot,
// data field unknown, item not in the hierarchy, partial / odd-numbered
// field arguments) the surface is `#REF!`, mirroring Mac Excel 365's
// behaviour.
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// dispatch-table contract in `tree_walker.cpp`.

#ifndef FORMULON_EVAL_GETPIVOTDATA_LAZY_H_
#define FORMULON_EVAL_GETPIVOTDATA_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `GETPIVOTDATA(data_field, pivot_anchor, [field1, item1, ...])` --
/// returns the value of a single pivot cell. `data_field` and the
/// optional field/item arguments are evaluated eagerly and coerced to
/// text; `pivot_anchor` is required to be a single cell `Ref` or a
/// `RangeOp` (whose top-left cell is used as the anchor). Any other AST
/// shape, an unknown data field, an unknown item, an arity mismatch
/// (no anchor, or odd field/item count), or an anchor not contained in
/// any pivot table on the workbook surfaces `#REF!`. Errors in any
/// argument propagate.
Value eval_getpivotdata_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_GETPIVOTDATA_LAZY_H_
