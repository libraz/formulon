// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Shared helpers for normalising the `Cell::formula_text` representation
// into the form the parser expects (a bare expression, no leading `=`).
//
// `Sheet::set_cell_formula` stores whatever string the caller hands it
// verbatim — most call sites strip the leading `=` before storing, but the
// reader / writer / external callers occasionally leave it in. The parser
// expects an expression (`1+2`), not an assignment (`=1+2`), so every
// recalc / resolve path that re-parses `formula_text` must strip the
// prefix first. This header centralises that strip so the various
// evaluator entry points cannot drift.

#ifndef FORMULON_EVAL_FORMULA_TEXT_UTILS_H_
#define FORMULON_EVAL_FORMULA_TEXT_UTILS_H_

#include <string_view>

namespace formulon {
namespace eval {

/// Returns `s` with a single leading '=' removed, if present.
///
/// The strip is intentionally one-shot (no loop) because Excel formulas
/// never legitimately begin with multiple `=` characters: `==1` would be
/// a parse error in Excel itself, and surfacing a parse error inside the
/// evaluator instead of silently coercing matches the source-of-truth
/// shape. A bare empty string is returned unchanged.
inline std::string_view strip_formula_prefix(std::string_view s) noexcept {
  if (!s.empty() && s.front() == '=') {
    s.remove_prefix(1);
  }
  return s;
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_FORMULA_TEXT_UTILS_H_
