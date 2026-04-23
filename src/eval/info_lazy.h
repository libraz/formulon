// Copyright 2026 libraz. Licensed under the MIT License.
//
// Lazy impls for the context-aware information predicates:
// ISFORMULA, ISREF, SHEET, SHEETS. Each inspects the un-evaluated AST
// of its argument (and, for SHEET / SHEETS, the bound Workbook /
// current Sheet on the EvalContext) rather than routing through a
// flattened `Value`. The eager dispatch path would discard the
// information they need — ISFORMULA looks up the referenced cell's
// `formula_text` on the Sheet, ISREF branches on the AST node kind,
// and SHEET / SHEETS consult the workbook geometry.
//
// See `eval/lazy_impls.h` for the shared `LazyImpl` signature and the
// dispatch-table contract in `tree_walker.cpp`.

#ifndef FORMULON_EVAL_INFO_LAZY_H_
#define FORMULON_EVAL_INFO_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `ISFORMULA(reference)` — returns TRUE iff `reference` is a literal
/// single-cell reference (NodeKind::Ref) pointing at a cell whose
/// `formula_text` is non-empty. Any other AST shape (literal, range,
/// function call, arithmetic) surfaces `#VALUE!`, matching Excel's
/// "ISFORMULA requires a reference" rule.
Value eval_isformula_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx);

/// `ISREF(value)` — returns TRUE iff the argument AST is a reference
/// shape (Ref / RangeOp / ExternalRef / StructuredRef / NameRef) or a
/// reference-returning call (INDIRECT / OFFSET / INDEX / CHOOSE) that
/// evaluates without error. Returns FALSE for literals, arithmetic,
/// and calls that produce a scalar Value. The predicate is lazy so a
/// static AST can answer without evaluation; reference-returning
/// calls are evaluated because their result depends on runtime input.
Value eval_isref_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);

/// `SHEET([value])` — returns the 1-based sheet number.
///   * 0 args: the currently-bound sheet's index + 1.
///   * 1 arg, a Ref / RangeOp with a sheet qualifier: the qualifier's
///     sheet number in the workbook.
///   * 1 arg, a Ref / RangeOp without a qualifier: current sheet.
///   * 1 arg evaluating to text: workbook.sheet_by_name(text).
///   * Missing sheet: `#N/A`. Unbound context: `#VALUE!`.
Value eval_sheet_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx);

/// `SHEETS([reference])` — returns the count of sheets in the
/// reference (3D reference width). This MVP lacks 3D references, so:
///   * 0 args: `workbook.sheet_count()` when bound, else 1.
///   * 1 arg that is a valid reference AST: 1.
///   * 1 arg that evaluates to an error: that error.
///   * 1 arg that is not a reference: `#VALUE!`.
Value eval_sheets_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_INFO_LAZY_H_
