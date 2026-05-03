// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `LambdaValue` is the closure backing-type for `Value::Lambda`. It is the
// runtime payload produced by evaluating a `parser::NodeKind::Lambda` AST
// node and consumed by `parser::NodeKind::LambdaCall`.
//
// Storage model: every member is a non-owning pointer into arena-allocated
// storage. The `params` array and the `LambdaValue` itself live in the
// evaluation arena; `body` references a parser-arena AST node; `captured_env`
// references a `NameEnv` frame chain that lives in the same evaluation arena
// as the surrounding `LET` (or is null at top-level scope). All four pointers
// must outlive any `Value::Lambda` that references them, the same lifetime
// contract as `Value::Text` / `Value::Array`.
//
// `LambdaValue` is intentionally pointer-only so the struct stays trivially
// copyable; the evaluator passes `Value` (which carries a single
// `const LambdaValue*`) through its value stack without heap allocation.
//
// This header sits in `eval/` rather than at the project root because it
// embeds eval-layer types (`NameEnv`) and parser AST nodes, neither of which
// the public `value.h` may depend on.

#ifndef FORMULON_EVAL_LAMBDA_VALUE_H_
#define FORMULON_EVAL_LAMBDA_VALUE_H_

#include <cstdint>
#include <string_view>
#include <type_traits>

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class NameEnv;

/// Runtime closure for an Excel `LAMBDA` form.
///
/// Built by the evaluator when it walks a `Lambda` AST node and consumed by
/// `LambdaCall`. The `captured_env` slot lets a lambda close over the
/// `LET`-bound names that were in scope at construction time, enabling
/// `=LET(y, 100, LAMBDA(x, x+y))(5)` to evaluate to 105 even though `y` is
/// out of lexical scope at the call site.
struct LambdaValue {
  /// Arena-allocated array of parameter names, length `param_count`.
  /// `nullptr` is legal when `param_count == 0`.
  const std::string_view* params;
  /// Number of declared parameters. `LambdaCall` arity must satisfy
  /// `param_count - optional_count <= arity <= param_count`; mismatches
  /// surface `#VALUE!`.
  std::uint32_t param_count;
  /// Number of trailing parameters declared with `[name]` bracket syntax.
  /// When the call site provides fewer than `param_count` arguments, the
  /// missing trailing slots are bound to an "omitted" sentinel that
  /// `ISOMITTED` detects.
  std::uint32_t optional_count;
  /// AST node to evaluate when the lambda is called. Non-null. Lifetime is
  /// bounded by the parser arena that produced the surrounding formula.
  const parser::AstNode* body;
  /// Lexical environment captured at construction time, or `nullptr` when no
  /// outer `LET` was in scope. Lifetime is bounded by the evaluation arena
  /// of the enclosing call to `evaluate()`.
  const NameEnv* captured_env;
};

// Trivially-copyable invariant: `Value` carries a `const LambdaValue*` and
// is itself trivially copyable; the pointed-to storage must therefore be
// trivially destructible (the arena never invokes destructors).
static_assert(std::is_trivially_destructible_v<LambdaValue>,
              "LambdaValue must be trivially destructible to live in Arena::create");
static_assert(std::is_trivially_copyable_v<LambdaValue>, "LambdaValue must be trivially copyable");

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LAMBDA_VALUE_H_
