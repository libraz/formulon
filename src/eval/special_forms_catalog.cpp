// Copyright 2026 libraz. Licensed under the MIT License.
//
// Sole source of truth for the parser-integrated special-form name list.
// See `eval/special_forms_catalog.h` for the rationale.

#include "eval/special_forms_catalog.h"

namespace formulon {
namespace eval {

const char* const* parser_special_form_names() {
  // Nullptr-terminated so callers can walk the array without a separate
  // length accessor. `LET` is intercepted by `parser::Parser` and lowered
  // to `NodeKind::Let`; it never reaches the `FunctionRegistry` or the
  // tree walker's `kLazyDispatch` table. Add future parser-only forms
  // (e.g. `LAMBDA`) here so the drift-detection pipeline picks them up
  // automatically.
  static constexpr const char* kNames[] = {"LET", nullptr};
  return kNames;
}

}  // namespace eval
}  // namespace formulon
