// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Registers Excel's BAHTTEXT function (Thai-baht spell-out) into a
// FunctionRegistry. Lives in its own translation unit because the algorithm
// is self-contained: a Thai-numeral spell-out does not reuse any of the
// helpers under `text.cpp` / `text_format.cpp`, and isolating it keeps the
// per-builtin object size small.

#ifndef FORMULON_EVAL_BUILTINS_TEXT_BAHTTEXT_H_
#define FORMULON_EVAL_BUILTINS_TEXT_BAHTTEXT_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers BAHTTEXT into `registry`. Intended to be invoked from
/// `register_builtins`.
void register_bahttext_builtin(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_TEXT_BAHTTEXT_H_
