//
// Identifies a syntactically omitted function argument (`f(a,,c)`). This is
// distinct from an evaluated blank reference: omission carries the callee's
// documented default rather than ordinary blank coercion.

#ifndef FORMULON_EVAL_OMITTED_ARG_H_
#define FORMULON_EVAL_OMITTED_ARG_H_

#include "parser/ast.h"

namespace formulon {
namespace eval {

inline bool is_omitted_arg(const parser::AstNode& arg) {
  return arg.kind() == parser::NodeKind::Literal && arg.as_literal().is_blank();
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_OMITTED_ARG_H_
