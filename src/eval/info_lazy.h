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

/// `FORMULATEXT(reference)` — returns the raw formula source string of
/// the referenced single-cell, including the leading `=`. The argument
/// must be a literal single-cell reference; anything else (literal,
/// RangeOp, non-reference call) surfaces `#VALUE!`. When the cell
/// carries no formula (empty `formula_text`, or the cell slot is
/// absent) the result is `#N/A`, matching Excel 365. A broken sheet
/// qualifier propagates as `#REF!` / `#NAME?` exactly like ISFORMULA.
Value eval_formulatext_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx);

/// `ISREF(value)` — returns TRUE iff the argument AST is a reference
/// shape (Ref / RangeOp / StructuredRef / NameRef) or a
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

/// `SINGLE(value)` — the explicit-name form of the `@` implicit
/// intersection operator. xlsx serialises `@A1:A5` as
/// `_xlfn.SINGLE(A1:A5)` so older Excel versions don't accidentally
/// expand it; the `_xlfn.` prefix is stripped before dispatch reaches
/// the registry. For a single-column range the formula cell's row
/// projects onto the range; for a single-row range the formula cell's
/// column projects. Other shapes surface `#VALUE!` (matching the `@`
/// operator's documented behaviour for 2-D and out-of-axis arguments).
Value eval_single_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx);

/// `CODE(text)` — routed lazily only so profile-specific codepage
/// fallbacks can see `EvalContext::excel_profile()`.
Value eval_code_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);
Value eval_lenb_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);
Value eval_weeknum_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx);
Value eval_text_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_INFO_LAZY_H_
