//
// Lazy impl entry point for `OFFSET(reference, rows, cols, [height],
// [width])`.

#ifndef FORMULON_EVAL_REFERENCE_OFFSET_H_
#define FORMULON_EVAL_REFERENCE_OFFSET_H_

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

/// `OFFSET(reference, rows, cols, [height], [width])` — offsets `reference`
/// by `(rows, cols)`, resizes it to `height` x `width`, and returns the
/// rectangle as a `Ref` marked dynamic, so the dereference that reads it
/// records the read.
Value eval_offset_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

// Compile-time guard: the lazy impl declared above must convert
// implicitly to the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`.
inline constexpr LazyImpl kOffsetLazySignatureWitness = &eval_offset_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_REFERENCE_OFFSET_H_
