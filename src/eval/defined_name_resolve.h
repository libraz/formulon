//
// Shared defined-name resolution.
//
// A workbook / sheet-scoped defined name (`Rate = 0.1`, `Sheet1!Local =
// Sheet1!$A$1`) must resolve identically for two consumers:
//
//   * the dependency extractor (`dep_extractor.cpp`), which walks a name's
//     body to collect the cells it reads; and
//   * the tree-walk evaluator (`tree_walker/walker.cpp`), which parses and
//     evaluates a name's body to a `Value`.
//
// The scope-priority lookup (sheet-scoped definition for the current sheet
// wins over a workbook-scoped one, case-insensitive) is identical for both,
// so it lives here as `find_defined_name`. The evaluation half
// (`resolve_defined_name`) is evaluator-only; the extractor keeps its own
// dep-collecting expansion.

#ifndef FORMULON_EVAL_DEFINED_NAME_RESOLVE_H_
#define FORMULON_EVAL_DEFINED_NAME_RESOLVE_H_

#include <cstdint>
#include <string_view>

#include "utils/arena.h"
#include "value.h"

namespace formulon {

class Workbook;
struct DefinedName;

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// Intrusive frame for defined-name cycle detection during evaluation.
///
/// Each active `resolve_defined_name` links one frame onto the chain carried
/// by `EvalContext` (see `EvalContext::with_defined_name_frame`). A self- or
/// mutually-referential name (`Loop = Loop + 1`, `A = B` / `B = A`) is then
/// detected by scanning the chain rather than recursing until the native
/// stack overflows. Frames live on the C++ call stack of the resolver, so the
/// chain is valid only for the duration of the resolution that built it.
struct DefinedNameFrame {
  /// Definition currently being expanded, owned by the workbook, which
  /// outlives the evaluation. Compared by identity: a workbook-scoped and a
  /// sheet-local name spelling the same text are different definitions.
  const DefinedName* definition = nullptr;
  /// Next frame further down the resolution stack, or null at the root.
  const DefinedNameFrame* prev = nullptr;
};

/// Finds the defined name `name` visible from sheet index `current_sheet_id`.
///
/// A sheet-scoped definition bound to `current_sheet_id` wins over a
/// workbook-scoped definition of the same name; matching is ASCII
/// case-insensitive, mirroring Excel's name-resolution semantics. Returns
/// `nullptr` when no definition matches.
const DefinedName* find_defined_name(const Workbook& workbook, std::uint16_t current_sheet_id,
                                     std::string_view name) noexcept;

/// Finds the defined name visible from an evaluator context. This is the
/// single context-aware lookup used by both ordinary NameRef evaluation and
/// named-LAMBDA call dispatch. The name is seen from
/// `ctx.name_scope_sheet()` when set, else from the current sheet; it
/// returns nullptr when the context is unbound, its current sheet is not
/// owned by the workbook, or no definition is visible.
const DefinedName* find_defined_name(const EvalContext& ctx, std::string_view name) noexcept;

/// Finds the definition a self-book reference `[0]!name` denotes: the
/// workbook-scoped one, else the sheet-local one on the lowest-index sheet,
/// as measured on Excel 365. Returns `nullptr` when `name` is undefined.
const DefinedName* find_self_book_defined_name(const Workbook& workbook, std::string_view name) noexcept;

/// Finds the defined name a sheet-qualified reference `sheet!name` denotes:
/// `name` as seen from `sheet`'s scope, so that sheet's local definition
/// wins over a workbook-scoped one. Returns `nullptr` when `sheet` names no
/// sheet of `workbook` or no definition is visible from it.
const DefinedName* find_sheet_defined_name(const Workbook& workbook, std::string_view sheet,
                                           std::string_view name) noexcept;

/// Parses the body of the already-located definition `def` into `arena` and
/// writes the context it evaluates in to `*out_ctx`: the using formula's
/// lexical scope cleared, `*frame` pushed onto the cycle chain, and a
/// sheet-local definition's own sheet as the name scope. Returns null with
/// `*out_err` set when `def` is null or its body is empty or unparsable
/// (`#NAME?`), or `def` is already being expanded (`#REF!`). `*frame` must
/// outlive every use of `*out_ctx`.
const parser::AstNode* prepare_defined_name_body(const DefinedName* def, Arena& arena, const EvalContext& ctx,
                                                 DefinedNameFrame* frame, EvalContext* out_ctx, ErrorCode* out_err);

/// Resolves the defined name `name` by parsing and evaluating its definition
/// in `ctx`. The definition may be a constant (`=0.1`), a reference
/// (`=Sheet1!$A$1`), or an arbitrary formula (`=A1*2`).
///
/// Returns:
///   * the evaluated `Value` on success;
///   * `#NAME?` when the name is undefined in scope, the context is unbound,
///     or the definition fails to parse;
///   * `#REF!` when a cycle is detected, matching the cell-cycle policy in
///     `EvalContext::resolve_ref`.
///
/// `arena` backs the parsed body and any text payload in the result; it must
/// outlive the returned `Value`. The definition is evaluated with the
/// caller's lexical `name_env()` cleared (a defined name is a top-level
/// formula and does not see the using formula's LET / LAMBDA bindings). A
/// sheet-local definition's body resolves its unqualified names in its own
/// sheet's scope; a workbook-scoped one inherits the caller's.
Value resolve_defined_name(std::string_view name, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx);

/// `resolve_defined_name` for the self-book spelling `[0]!name` (see
/// `find_self_book_defined_name`). An undefined name yields `#NAME?`.
Value resolve_self_book_defined_name(std::string_view name, Arena& arena, const FunctionRegistry& registry,
                                     const EvalContext& ctx);

/// `resolve_defined_name` for the sheet-qualified spelling `sheet!name`
/// (see `find_sheet_defined_name`). A `sheet` naming no sheet resolves as
/// the book-scope name of the linked book `Workbook::link_for_sheet_qualifier`
/// finds, else yields `#REF!` as `NoSuchSheet!A1` does; an undefined name
/// yields `#NAME?`.
Value resolve_sheet_defined_name(std::string_view sheet, std::string_view name, Arena& arena,
                                 const FunctionRegistry& registry, const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DEFINED_NAME_RESOLVE_H_
