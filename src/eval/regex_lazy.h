// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Lazy impls for the Excel REGEX function family backed by PCRE2:
//   * REGEXTEST(text, pattern, [case_sensitivity]) -> bool
//   * REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity]) ->
//     scalar text or array of matches / capture groups
//   * REGEXREPLACE(text, pattern, replacement, [occurrence],
//     [case_sensitivity]) -> text
//
// They ride the lazy seam because their `text` argument may be a
// Range / Ref / ArrayLiteral that must reach the impl with its (rows,
// cols) shape preserved (matches Mac Excel: a range hands back an
// array-shaped result, broadcasting cellwise).
//
// Compile + match runs through a single shared kernel in
// `regex_lazy.cpp`; see that file for the option set, error mapping,
// and resource limits applied to every regex evaluation.

#ifndef FORMULON_EVAL_REGEX_LAZY_H_
#define FORMULON_EVAL_REGEX_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `REGEXTEST(text, pattern, [case_sensitivity=0])` -> Bool. Returns TRUE
/// when `pattern` matches anywhere in `text`. `case_sensitivity` must
/// coerce to {0, 1}; any other value yields `#VALUE!`. The convention
/// matches Mac Excel 365 / MS docs: `0`/FALSE (default) is
/// case-sensitive, `1`/TRUE is case-insensitive. Resource exhaustion
/// (match_limit / depth_limit) returns FALSE rather than `#CALC!`
/// because the test is a predicate by intent.
Value eval_regextest_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx);

/// `REGEXEXTRACT(text, pattern, [return_mode=0], [case_sensitivity=0])`.
/// `return_mode` selects the result shape:
///   * 0 -> scalar text of the first whole match.
///   * 1 -> row array (1 x N) of all whole matches.
///   * 2 -> row array (1 x G) of capture groups from the FIRST match.
///   * 3 -> 2D array (N x G) of capture groups from every match.
/// Modes 2 and 3 yield `#VALUE!` when the pattern has no capture groups
/// AND a match occurred (no match still yields `#N/A`). Resource
/// exhaustion yields `#CALC!`.
Value eval_regexextract_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx);

/// `REGEXREPLACE(text, pattern, replacement, [occurrence=0],
/// [case_sensitivity=0])`. `occurrence == 0` replaces every match
/// (PCRE2_SUBSTITUTE_GLOBAL); `occurrence == N > 0` replaces only the
/// N-th match. Mac Excel 365 is permissive on negative values: they
/// fold to global replacement. When `occurrence` exceeds the number of
/// matches the original `text` is returned unchanged. Resource
/// exhaustion yields `#CALC!`.
Value eval_regexreplace_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_REGEX_LAZY_H_
