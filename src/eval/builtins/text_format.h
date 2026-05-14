// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Registers Excel's text-conversion built-ins (TEXT, VALUE, NUMBERVALUE)
// into a FunctionRegistry. Kept in its own translation unit because the
// three builtins share the format-string engine under
// `eval/text_format/number_format.{h,cpp}` and the date/time parser under
// `eval/date_text_parse.{h,cpp}`.

#ifndef FORMULON_EVAL_BUILTINS_TEXT_FORMAT_H_
#define FORMULON_EVAL_BUILTINS_TEXT_FORMAT_H_

#include "eval/lazy_impls.h"

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers TEXT, VALUE, and NUMBERVALUE into `registry`. Intended to be
/// invoked from `register_builtins`.
void register_text_format_builtins(FunctionRegistry& registry);

/// ARRAYTOTEXT(array, [format]) must preserve the 2-D shape of range and
/// inline-array arguments, so it rides the lazy dispatch path.
Value eval_arraytotext_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_TEXT_FORMAT_H_
