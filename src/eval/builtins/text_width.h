// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers the Excel width-conversion text built-ins: ASC (full-width to
// half-width) and JIS / DBCS (half-width to full-width). The mapping tables
// are hard-coded against the ja-JP Excel 365 observed behaviour; see the
// implementation file for the per-codepoint contract and notes on hiragana /
// archaic katakana passthrough.

#ifndef FORMULON_EVAL_BUILTINS_TEXT_WIDTH_H_
#define FORMULON_EVAL_BUILTINS_TEXT_WIDTH_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the width-conversion text built-ins (ASC, JIS, DBCS) into
/// `registry`. Intended to be invoked from `register_builtins`.
void register_text_width_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_TEXT_WIDTH_H_
