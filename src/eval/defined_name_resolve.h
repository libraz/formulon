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

namespace io {
struct DefinedName;
}  // namespace io

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
  /// Authored name currently being expanded. Borrows the `DefinedName::name`
  /// string owned by the workbook, which outlives the evaluation.
  std::string_view name;
  /// Next frame further down the resolution stack, or null at the root.
  const DefinedNameFrame* prev = nullptr;
};

/// Finds the defined name `name` visible from sheet index `current_sheet_id`.
///
/// A sheet-scoped definition bound to `current_sheet_id` wins over a
/// workbook-scoped definition of the same name; matching is ASCII
/// case-insensitive, mirroring Excel's name-resolution semantics. Returns
/// `nullptr` when no definition matches.
const io::DefinedName* find_defined_name(const Workbook& workbook, std::uint16_t current_sheet_id,
                                         std::string_view name) noexcept;

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
/// formula and does not see the using formula's LET / LAMBDA bindings).
Value resolve_defined_name(std::string_view name, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DEFINED_NAME_RESOLVE_H_
