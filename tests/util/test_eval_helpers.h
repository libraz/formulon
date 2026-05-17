// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Test-side evaluator wrappers.
//
// Mirror the open-coded `EvalSource` / `EvalSourceIn` / `EvalSourceAt`
// helpers currently re-defined in 40+ test files. Each wrapper parses
// `formula` with `parser::Parser`, then invokes the production tree
// walker (`eval::evaluate`) and returns the resulting `Value`.
//
// Arenas come from `test_arena.h`. Each call resets the relevant arena
// at entry so consecutive invocations do not accumulate AST or runtime
// payloads. Callers that need to hold a Value across two consecutive
// helper calls (e.g. comparing the result of two formulas) should copy
// any text payload out before issuing the second call - the underlying
// arena bytes are reused.
//
// IMPORTANT: this is NOT the same header as
// `tests/unit/eval/test_eval_helpers.h`, which exposes profile-aware
// EvalContext factories (`mac_context()`, `win_context()`). The two are
// distinct on purpose: profile selection is a per-test policy decision
// and is kept visible at call sites; this header centralises only the
// mechanical parse-then-evaluate boilerplate.

#ifndef FORMULON_TESTS_UTIL_TEST_EVAL_HELPERS_H_
#define FORMULON_TESTS_UTIL_TEST_EVAL_HELPERS_H_

#include <cstdint>
#include <string_view>

#include "util/test_arena.h"
#include "value.h"

namespace formulon {

class Sheet;
class Workbook;

namespace test {

// Parses `formula` and evaluates it through the default function
// registry with no bound workbook. Equivalent to the open-coded
// `EvalSource(std::string_view)` helper duplicated across
// `tests/unit/eval/builtins_*_test.cpp`.
//
// Parse failures (root == nullptr) surface as `#NAME?`; this matches
// the duplicated helpers' fallback behaviour and keeps test diagnostics
// uniform.
Value EvalSource(std::string_view formula);

// Parses `formula` and evaluates it against a bound workbook +
// current sheet, using the default function registry. Used by tests
// that exercise reference / range resolution. A fresh `EvalState` is
// constructed internally; tests that need to inspect or share state
// across helper calls should drop down to the explicit `evaluate(...)`
// overload directly.
Value EvalSourceIn(std::string_view formula, const Workbook& wb, const Sheet& current);

// Same as `EvalSourceIn` but anchors the formula cell at (`row`,
// `col`) via `EvalContext::with_formula_cell`. Needed for CELL,
// implicit-intersection, and other anchor-sensitive functions.
Value EvalSourceAt(std::string_view formula, const Workbook& wb, const Sheet& current, std::uint32_t row,
                   std::uint32_t col);

// Convenience overload: looks up the current sheet by name in `wb`.
// Returns `#REF!` when the sheet name does not match any sheet in the
// workbook. Mirrors the spec-level signature in
// `docs/refactor/03-common-modules.md` C-11 without forcing callers
// who already have a `Sheet&` to re-resolve it by name.
Value EvalSourceAt(std::string_view formula, const Workbook& wb, std::string_view sheet, std::uint32_t row,
                   std::uint32_t col);

}  // namespace test
}  // namespace formulon

#endif  // FORMULON_TESTS_UTIL_TEST_EVAL_HELPERS_H_
