//
// Internal seam between the recursive node visitor (`tree_walker/walker.cpp`),
// the bare-range / intersection spill positions
// (`tree_walker/spill_position.cpp`) and the top-level evaluation entry
// (`tree_walker/evaluate_top.cpp`). Not part of the public evaluator surface.

#ifndef FORMULON_EVAL_TREE_WALKER_WALKER_INTERNAL_H_
#define FORMULON_EVAL_TREE_WALKER_WALKER_INTERNAL_H_

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tail_array.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

// Materialises the inclusive rectangle [`top_left`, `bottom_right`] as a
// `Value::Array`. Both endpoints must be bounded (no `is_full_col` /
// `is_full_row`) and already normalised so that `top_left` is the smaller
// corner on both axes. Cells the expansion did not produce are padded with
// Blank, which the top-level surface contract later projects to 0.
Value materialize_rectangle(const parser::Reference& top_left, const parser::Reference& bottom_right, Arena& arena,
                            const FunctionRegistry& registry, const EvalContext& ctx);

// Recovers the rectangle a bare whole-axis reference declares; false (out-params
// untouched) when `node` is not one.
bool whole_axis_declared_rect(const parser::AstNode& node, parser::Reference* top_left,
                              parser::Reference* bottom_right);

// Recovers the rectangle a bare bounded range declares (`A1:C10`); false
// (out-params untouched) for any other shape.
bool bounded_declared_rect(const parser::AstNode& node, parser::Reference* top_left, parser::Reference* bottom_right);

// Reads the rectangle `[lhs : rhs]`, which spans a whole grid axis, at its
// declared size as a `TailArray`.
Shaped read_whole_axis(const parser::Reference& lhs, const parser::Reference& rhs, Arena& arena,
                       const FunctionRegistry& registry, const EvalContext& ctx);

// How a bare range standing as the formula is evaluated: whether a range
// holding the anchor settles as a cycle, and whether a whole-axis rectangle may
// come back as a `TailArray`.
struct BareRangeSpillOptions {
  bool settle_circularity = false;
  bool allow_tail = false;
};

// Evaluates a bare range standing as the entire formula, weighing it against
// the anchor's spill footprint. `*footprint_probed` is set when it was measured.
Shaped evaluate_bare_range_spill(const parser::Reference& top_left, const parser::Reference& bottom_right, Arena& arena,
                                 const FunctionRegistry& registry, const EvalContext& ctx,
                                 const BareRangeSpillOptions& options, bool* footprint_probed);

// Evaluates the space-as-intersection operator; `spill_position` marks the
// whole-formula position.
Value eval_intersect_op(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, bool spill_position);

// First element of a shaped result, read without expanding it.
Value first_element(const Shaped& s);

// The value of a formula's root node.
Shaped eval_root_shaped(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TREE_WALKER_WALKER_INTERNAL_H_
