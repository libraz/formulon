//
// Lazy impl for `_xlfn.ANCHORARRAY(ref)` — the OOXML internal encoding
// of the postfix spill operator `ref#`. Excel itself transparently strips
// the `_xlfn.` prefix on load, so callers see plain `ANCHORARRAY(ref)`
// at the dispatch layer. Returns the spill region anchored at `ref` as a
// `Value::Array`, mirroring the `NodeKind::SpillRef` branch in
// `tree_walker.cpp`.
//
// It is lazy (rather than eager like ordinary builtins) because the
// argument must reach the impl as an AST node so a bare `Ref` / `SpillRef`
// can be inspected without being flattened to a scalar `Value`, and so a
// reference-returning `Call` (OFFSET / INDIRECT) can be resolved with its
// rectangle metadata intact. The central dispatch table in
// `tree_walker_lazy_table.cpp` references the extern by unqualified name;
// see `eval/lazy_impls.h` for the shared `LazyImpl` signature.

#ifndef FORMULON_EVAL_DYNAMIC_ARRAY_ANCHOR_H_
#define FORMULON_EVAL_DYNAMIC_ARRAY_ANCHOR_H_

#include "eval/lazy_impls.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

Value eval_anchorarray_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx);

// Compile-time guard: the lazy impl declared above must convert
// implicitly to the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`.
inline constexpr LazyImpl kDynamicArrayAnchorLazySignatureWitness = &eval_anchorarray_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DYNAMIC_ARRAY_ANCHOR_H_
