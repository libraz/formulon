//
// TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...) — joins every
// text argument with `delimiter`, skipping empty pieces when
// `ignore_empty` is TRUE.
//
// A lazy (AST-driven) impl rather than an eager, range-flattening
// FunctionDef entry: the generic `accepts_ranges` flatten in the
// tree-walker's dispatcher expands every positional argument uniformly by
// row-major cell order, which corrupts TEXTJOIN's two fixed leading
// positions (`delimiter`, `ignore_empty`) whenever either is itself a
// range or array — Excel applies a range/array delimiter cyclically
// across the flattened text arguments, not spliced into the argument
// list by position. Routing TEXTJOIN through the AST directly keeps that
// distinction without touching the shared dispatcher every other
// range-aware aggregator relies on.

#ifndef FORMULON_EVAL_TEXTJOIN_LAZY_H_
#define FORMULON_EVAL_TEXTJOIN_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)`. `delimiter`
/// may be a scalar or a range/array; an array delimiter is applied
/// cyclically across the flattened, in-call-order sequence of `text1,
/// text2, ...` cells (each of which may itself be a scalar, range, or
/// array). Errors in any argument propagate. Result length is capped at
/// Excel's 32,767 UTF-16-unit limit; exceeding it surfaces `#CALC!`, counting
/// the blank cells of a whole-axis reference past its populated head.
Value eval_textjoin_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXTJOIN_LAZY_H_
