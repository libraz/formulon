// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Renders an `eval::LambdaValue` (the runtime closure produced by
// evaluating a `=LAMBDA(...)` form) back into Excel formula text.
//
// `LambdaValue` carries the parameter list and the body AST node
// separately; reconstructing the surface `LAMBDA(p1,p2,body)` form is
// non-trivial because the body itself is a generic AST subtree (which
// `parser::format_formula` already handles) but the parameter list
// and bracket-syntax for optional trailing parameters are owned by
// the runtime closure, not the AST.
//
// This helper sits in `eval/` rather than `parser/` because the parser
// layer must not depend on `eval::LambdaValue`.

#ifndef FORMULON_EVAL_LAMBDA_FORMAT_H_
#define FORMULON_EVAL_LAMBDA_FORMAT_H_

#include <string>

namespace formulon {
namespace eval {

struct LambdaValue;

/// Returns the textual `LAMBDA(params..., body)` rendering of `lv`.
///
/// The returned string never carries a leading `=`. Optional trailing
/// parameters (the last `lv.optional_count` slots of `lv.params`) are
/// emitted with bracket syntax, matching the source surface that
/// `parser::Parser` accepts. The body is rendered via
/// `parser::format_formula`, so the same round-trip equivalence
/// contract applies: parsing the returned text yields a structurally
/// equivalent lambda AST.
std::string format_lambda_value(const LambdaValue& lv);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LAMBDA_FORMAT_H_
