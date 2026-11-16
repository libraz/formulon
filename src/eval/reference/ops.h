//
// The reference operators over `Ref` values: `:` (bounding box), space
// (intersection) and `,` (union). Each takes the operands as `Ref` values and
// returns a `Ref` or the error `Value`; an error operand is returned as it is.
// The result is `kDynamic` when any operand is.
//
// Sheet qualifiers merge by `RefOrigin`: a bare (`kHomeSheet`) operand
// inherits the other operand's sheet; a `kDerived` sheet (one a name or call
// produced) does not reach a bare operand unless it is the formula's own
// sheet. The result's origin is `kSpelled` if either operand is, else
// `kDerived` if either is, else `kImplicit`.

#ifndef FORMULON_EVAL_REFERENCE_OPS_H_
#define FORMULON_EVAL_REFERENCE_OPS_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

class EvalContext;

/// `lhs : rhs`, the bounding box of two `kArea` operands. Chained use gives
/// the box of every endpoint. A non-`Ref`, union, 3-D or grid operand, two
/// different sheets, or two different books give `#VALUE!`.
Value ref_range(const Value& lhs, const Value& rhs, Arena& arena, const EvalContext& ctx);

/// `lhs rhs`, the cells both operands cover: a `kArea` for one area, a
/// `kUnion` for several, `#NULL!` for none. A union operand intersects area by
/// area. Two different explicit sheets give `#REF!`; a non-`Ref`, 3-D or grid
/// operand, or two different books, give `#VALUE!`.
Value ref_intersect(const Value& lhs, const Value& rhs, Arena& arena, const EvalContext& ctx);

/// `lhs , rhs`, the `kUnion` of both operands' areas in order. A 3-D operand
/// gives `#NAME?`. A non-`Ref` or grid operand, two different books, or
/// spelled areas on two different sheets give `#VALUE!`.
Value ref_union(const Value& lhs, const Value& rhs, Arena& arena, const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_REFERENCE_OPS_H_
