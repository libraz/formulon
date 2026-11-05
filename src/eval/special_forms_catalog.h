//
// Enumeration seam for parser-integrated special forms.
//
// A "parser-integrated special form" is a function-like name that Formulon
// recognises during parsing (the parser carves it out of the ordinary
// `Call` node path and lowers it to a dedicated AST shape) and therefore
// never reaches the `FunctionRegistry` or the tree walker's `kLazyDispatch`
// table. The list currently contains `LET` and `LAMBDA`.
//
// The drift-detection machinery (`tests/unit/registry_catalog_test.cpp`
// and `tools/catalog/status.py`) enumerates these names alongside the
// registry and lazy-dispatch tables so the catalog invariant includes
// parser-only forms too.

#ifndef FORMULON_EVAL_SPECIAL_FORMS_CATALOG_H_
#define FORMULON_EVAL_SPECIAL_FORMS_CATALOG_H_

#include <cstdint>
#include <string_view>

namespace formulon {
namespace eval {

/// Returns a nullptr-terminated static array of canonical UPPERCASE names
/// for forms that are recognised by the parser but NOT registered via
/// `FunctionRegistry` and NOT routed through `kLazyDispatch`: `LET` and
/// `LAMBDA`.
///
/// The returned array has program lifetime and must not be freed.
const char* const* parser_special_form_names();

/// Stores the inclusive argument-count bounds of the special form `name`
/// (ASCII case-insensitive) and returns true, or returns false when `name`
/// is not a special form.
bool find_special_form_arity(std::string_view name, std::uint32_t* min_arity, std::uint32_t* max_arity);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SPECIAL_FORMS_CATALOG_H_
