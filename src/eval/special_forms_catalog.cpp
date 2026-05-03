// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Sole source of truth for the parser-integrated special-form name list.
// See `eval/special_forms_catalog.h` for the rationale.

#include "eval/special_forms_catalog.h"

namespace formulon {
namespace eval {

const char* const* parser_special_form_names() {
  // Nullptr-terminated so callers can walk the array without a separate
  // length accessor. `LET` and `LAMBDA` are both lowered by the parser to
  // dedicated AST shapes (`NodeKind::LetBinding` and `NodeKind::Lambda` /
  // `NodeKind::LambdaCall`); neither name reaches the `FunctionRegistry`
  // or the tree walker's `kLazyDispatch` table.
  static constexpr const char* kNames[] = {"LET", "LAMBDA", nullptr};
  return kNames;
}

}  // namespace eval
}  // namespace formulon
