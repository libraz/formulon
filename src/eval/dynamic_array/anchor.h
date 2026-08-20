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

#include <cstdint>
#include <string_view>

#include "eval/lazy_impls.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// Resolves `node` to the cell a spill anchor names, writing the sheet
/// qualifier (empty = bound sheet) and 0-based coordinates to the out
/// parameters. Shared by `ANCHORARRAY` and by the `NodeKind::SpillRef`
/// branches, which are two spellings of one operator.
///
/// Accepts the shapes Excel accepts as an anchor: a written-out `Ref`, a
/// nested `SpillRef`, a reference-returning `Call` (OFFSET / INDIRECT /
/// IF / CHOOSE), and a `NameRef` bound to one of those by an enclosing
/// LET. An anchor is a single cell, so a reference that resolves to more
/// than one is rejected rather than narrowed to its top-left corner:
/// `=SUM(OFFSET(A1,0,0,2,2)#)` is `#REF!` in Excel even when A1 anchors a
/// live spill.
///
/// Returns false on failure with the Excel error code in `*out_err`.
/// Every anchor that fails to resolve surfaces `#REF!`, matching Excel,
/// which does not distinguish a missing spill from an unusable anchor.
bool resolve_spill_anchor_node(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_row,
                               std::uint32_t* out_col, ErrorCode* out_err);

Value eval_anchorarray_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx);

// Compile-time guard: the lazy impl declared above must convert
// implicitly to the shared `LazyImpl` function-pointer type published in
// `eval/lazy_impls.h`.
inline constexpr LazyImpl kDynamicArrayAnchorLazySignatureWitness = &eval_anchorarray_lazy;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DYNAMIC_ARRAY_ANCHOR_H_
