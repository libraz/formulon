//
// Lazy-dispatch seam between `tree_walker.cpp` and the per-family lazy-impl
// translation units.
//
// Why this header exists: the dispatch table itself (a constexpr array
// mapping canonical UPPERCASE function names to `LazyImpl` pointers) used
// to live at the top of `tree_walker.cpp`. The table forced
// `tree_walker.cpp` to `#include` every per-family header (30+ entries),
// which made the evaluator's main TU a hot rebuild target whenever a
// single new lazy function was added. By moving the table into its own
// TU (`tree_walker_lazy_table.cpp`) and exposing only the two narrow
// hooks below, adding a new lazy entry only rebuilds:
//   - the family TU that owns the new impl, and
//   - `tree_walker_lazy_table.cpp`.
// `tree_walker.cpp` itself is untouched.
//
// Both functions are stateless and thread-safe. `find_lazy_impl` is
// case-insensitive on ASCII (matches `case_insensitive_eq` semantics).

#ifndef FORMULON_EVAL_TREE_WALKER_LAZY_TABLE_H_
#define FORMULON_EVAL_TREE_WALKER_LAZY_TABLE_H_

#include <string_view>

#include "eval/lazy_impls.h"

namespace formulon {
namespace eval {

/// Looks up `name` (canonical UPPERCASE, case-insensitive) in the lazy
/// dispatch table. Returns the bound `LazyImpl` pointer on hit, or
/// `nullptr` when the name is not registered as a lazy form. Callers
/// invoke the returned pointer directly with the `Call` AST node + the
/// usual `Arena` / `FunctionRegistry` / `EvalContext` triple.
LazyImpl find_lazy_impl(std::string_view name) noexcept;

/// Returns a nullptr-terminated array of canonical UPPERCASE names in
/// the lazy dispatch table. Storage has static duration; callers must
/// not free the array. Used by `tree_walker.cpp::lazy_form_names`
/// (declared in `tree_walker.h`) to expose the routing decisions to
/// catalog/diagnostics consumers.
const char* const* lazy_table_names();

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TREE_WALKER_LAZY_TABLE_H_
