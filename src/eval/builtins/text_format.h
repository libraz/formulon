// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's text-conversion built-ins (TEXT, VALUE, NUMBERVALUE)
// into a FunctionRegistry. Kept in its own translation unit because the
// three builtins share the format-string engine under
// `eval/text_format/number_format.{h,cpp}` and the date/time parser under
// `eval/date_text_parse.{h,cpp}`.

#ifndef FORMULON_EVAL_BUILTINS_TEXT_FORMAT_H_
#define FORMULON_EVAL_BUILTINS_TEXT_FORMAT_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers TEXT, VALUE, and NUMBERVALUE into `registry`. Intended to be
/// invoked from `register_builtins`.
void register_text_format_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_TEXT_FORMAT_H_
