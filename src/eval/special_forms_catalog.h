// Copyright 2026 libraz. Licensed under the MIT License.
//
// Enumeration seam for parser-integrated special forms.
//
// A "parser-integrated special form" is a function-like name that Formulon
// recognises during parsing (the parser carves it out of the ordinary
// `Call` node path and lowers it to a dedicated AST shape) and therefore
// never reaches the `FunctionRegistry` or the tree walker's `kLazyDispatch`
// table. Today that is just `LET`; once `LAMBDA` lands it joins the list.
//
// The drift-detection machinery (`tests/unit/registry_catalog_test.cpp`
// and `tools/catalog/status.py`) enumerates these names alongside the
// registry and lazy-dispatch tables so that the catalog invariant and
// the coverage report stay accurate for parser-only forms too.

#ifndef FORMULON_EVAL_SPECIAL_FORMS_CATALOG_H_
#define FORMULON_EVAL_SPECIAL_FORMS_CATALOG_H_

namespace formulon {
namespace eval {

/// Returns a nullptr-terminated static array of canonical UPPERCASE names
/// for forms that are recognised by the parser but NOT registered via
/// `FunctionRegistry` and NOT routed through `kLazyDispatch`. Today: `LET`.
/// Future: `LAMBDA`.
///
/// The returned array has program lifetime and must not be freed.
const char* const* parser_special_form_names();

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SPECIAL_FORMS_CATALOG_H_
