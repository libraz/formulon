//
// Lazy impl entry point for `INDIRECT(ref_text, [a1])`. Declared in its
// own header so the central dispatch table (`tree_walker_lazy_table.cpp`)
// can include it directly without pulling in the OFFSET / intersection
// surfaces. The accompanying text-to-rectangle decoder lives in
// `eval/a1_parse.h` (public, defined in `eval/a1_parse.cpp`) and
// `reference/common.h` (intra-subdir).

#ifndef FORMULON_EVAL_REFERENCE_INDIRECT_H_
#define FORMULON_EVAL_REFERENCE_INDIRECT_H_

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

/// `INDIRECT(ref_text, [a1])` — evaluates `ref_text`, parses it as an A1 or
/// (`a1 = FALSE`) R1C1 reference, and returns that reference as a `Ref`
/// marked dynamic, so the dereference that reads it records the read. Empty
/// or malformed text, or a sheet the workbook does not hold, is `#REF!`.
Value eval_indirect_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);

// Compile-time guard: the lazy impl declared above must convert
// implicitly to the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`.
inline constexpr LazyImpl kIndirectLazySignatureWitness = &eval_indirect_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_REFERENCE_INDIRECT_H_
