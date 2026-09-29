//
// Lazy impls for the "classic" lookup family: `CHOOSE`, `INDEX`, `MATCH`,
// `VLOOKUP`, and `HLOOKUP`. These builtins share the legacy row/column
// scan model (linear walk along a chosen axis with approximate /
// exact-with-wildcard modes) and a single private helper, `lookup_scan`,
// which lives alongside the impls in `lookups/classic.cpp`.
//
// They are lazy (rather than eager like ordinary builtins) because the
// lookup array argument must reach the impl as an AST node so a bare
// single-cell `Ref` can still be treated as a 1-cell range, and because
// scalar-index `CHOOSE` must select exactly one of its value arguments and
// leave the rest un-evaluated. Array-index `CHOOSE` deliberately evaluates
// every branch once. The central dispatch table in `tree_walker.cpp`
// references these externs by unqualified name; see `eval/lazy_impls.h`
// for the shared `LazyImpl` signature and `eval_node` entry point.

#ifndef FORMULON_EVAL_LOOKUPS_CLASSIC_H_
#define FORMULON_EVAL_LOOKUPS_CLASSIC_H_

#include <cstdint>
#include <string_view>

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

Value eval_choose_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);
/// Evaluates an array-index CHOOSE after its index has already been
/// materialised. Every branch is evaluated exactly once, left-to-right, and
/// the selected cells are composed into one value array. Callers must pass an
/// Array-valued index; this seam is shared by the lazy evaluator and the
/// range-aware CHOOSE expander so the value-array blank marker is preserved.
Value eval_choose_array_index_lazy(const parser::AstNode& call, const Value& index_value, Arena& arena,
                                   const FunctionRegistry& registry, const EvalContext& ctx);
Value eval_index_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);
/// Resolves an `INDEX(reference, row_num, [col_num], [area_num])` call to
/// the rectangle it selects, without reading any cell: one cell, or a whole
/// row / column / area for a zero index. `area_num` picks from a
/// parenthesised union in source order. Fails with `#VALUE!` when the
/// selected area is not a reference or a selector is an array, and with the
/// value path's error otherwise. Writes the 0-based inclusive rectangle and
/// its sheet qualifier (empty = bound sheet).
bool resolve_index_reference(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                             std::uint32_t* out_left_col, std::uint32_t* out_bottom_row, std::uint32_t* out_right_col,
                             ErrorCode* out_err);
Value eval_match_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);
Value eval_vlookup_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);
Value eval_hlookup_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);
Value eval_lookup_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

// Compile-time guard: every lazy impl declared above must convert
// implicitly to the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`. Picking `eval_choose_lazy` as a witness is
// sufficient because every sibling shares the same parameter list.
inline constexpr LazyImpl kClassicLookupsLazySignatureWitness = &eval_choose_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LOOKUPS_CLASSIC_H_
