//
// Projection of a structured (table) reference onto the value model:
// rectangle resolution, the single-cell shortcut and the blank padding.

#ifndef FORMULON_EVAL_STRUCTURED_REF_PROJECT_H_
#define FORMULON_EVAL_STRUCTURED_REF_PROJECT_H_

#include <string_view>

#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

class EvalContext;
class FunctionRegistry;

// Resolves `table_name[column_payload]` to a concrete rectangle on the
// table's home sheet and reads it through the same expansion path a
// literal `Sheet1!A1:C10` takes, so cross-sheet resolution and cycle
// detection stay in one place. A single-cell rectangle reads as a plain
// reference; a larger one becomes a `Value::Array` allocated in `arena`,
// padded with Blank where the expansion returned fewer cells than the
// rectangle covers.
//
// Failures come back as an error `Value`. `*out_arena_exhausted` is set
// only when the array allocation itself failed — the tree walker
// surfaces that as the returned `#NUM!`, while the VM replaces it with
// its structured out-of-memory error.
Value project_structured_ref(std::string_view table_name, std::string_view column_payload, Arena& arena,
                             const FunctionRegistry& registry, const EvalContext& ctx, bool* out_arena_exhausted);

// The endpoint pair of the static reference `node` names, as
// `declared_rect_endpoint_pair` reports it, extended to structured
// references: `T[p]` yields the sheet-qualified corners of the rectangle it
// resolves to, so every reader treats it exactly as that A1 range. False for
// any other shape, and for a structured reference that does not resolve,
// whose error then goes to `*out_err` when given.
bool reference_endpoint_pair(const parser::AstNode& node, parser::Reference* out_lhs, parser::Reference* out_rhs,
                             const EvalContext& ctx, ErrorCode* out_err = nullptr);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_STRUCTURED_REF_PROJECT_H_
