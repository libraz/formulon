// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's text built-ins (UPPER, LOWER, TRIM, LEFT, RIGHT, MID,
// REPT, SUBSTITUTE, FIND, SEARCH, VALUE, EXACT, TEXTJOIN, UNICHAR, UNICODE,
// CLEAN, PROPER) into a FunctionRegistry. Kept in its own translation unit
// so the text family can evolve independently of the rest of the builtin
// catalog.

#ifndef FORMULON_EVAL_BUILTINS_TEXT_H_
#define FORMULON_EVAL_BUILTINS_TEXT_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the text built-in functions (UPPER, LOWER, TRIM, LEFT, RIGHT,
/// MID, REPT, SUBSTITUTE, FIND, SEARCH, VALUE, EXACT, TEXTJOIN, UNICHAR,
/// UNICODE, CLEAN, PROPER) into `registry`. Intended to be invoked from
/// `register_builtins`.
void register_text_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_TEXT_H_
